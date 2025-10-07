// Copyright 2022 exl Author. All Rights Reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//      http://www.apache.org/licenses/LICENSE-2.0
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

package exl

import (
	"encoding"
	"errors"
	"fmt"
	"io"
	"reflect"
	"time"

	"codeberg.org/tealeg/xlsx/v4"
)

type (
	ReadConfigurator             interface{ ReadConfigure(rc *ReadConfig) }
	RowUnmarshalErrorHandlerFunc func(*xlsx.Cell, *reflect.Value, FieldInfo)
	UnusedColumnsHandlerFunc     func(*xlsx.Cell, *reflect.Value, FieldInfo)
	ReadConfig                   struct {
		// The tag name to use when looking for fields in the target struct.
		// Defaults to "excel".
		TagName string
		// Name of the worksheet to be read. Takes precedence over SheetIndex.
		// Defaults to ""
		SheetName string
		// The index of the worksheet to be read.
		// Defaults to 0, the first worksheet.
		SheetIndex int
		// The row index at which the column headers are read from.
		// Zero-based, defaults to 0.
		HeaderRowIndex int
		// Start the data reading at this row.
		// The header row counts as row.
		// Zero-based, defaults to 1.
		DataStartRowIndex int
		// Configure the default string unmarshaler to trim space after reading a cell.
		// Does not impact any other default unmarshaler,
		// but is available to custom unmarshalers via ExcelUnmarshalParameters.TrimSpace.
		// Defaults to false.
		TrimSpace bool
		// Fallback date formats for date parsing.
		// If an Excel cell is to be unmarshalled into a date,
		// and that cell is either not formatted as Date or contains raw text
		// (which can happen if Excel does not correctly recognize the date format)
		// then these formats are used in the order specified to try and parse
		// the raw cell value into a date.
		// There are no fallback formats configured by default.
		FallbackDateFormats []string
		// Skip reading columns for which no target field is found.
		// Defaults to true.
		SkipUnknownColumns bool
		// Skip reading columns, if there is a target field,
		// but the target type is unsupported
		// or caused an error when determining the unmarshaler to use.
		// Defaults to false.
		SkipUnknownTypes bool
		// Configure how errors during unmarshalling are handled.
		// Unmarshalling errors are e.g. invalid number formats in the cell,
		// date parsing with invalid input,
		// or attempting to unmarshal non-numeric text into a numeric field.
		// Defaults to UnmarshalErrorAbort.
		UnmarshalErrorHandling UnmarshalErrorHandling
		// If UnmarshalErrorHandling is configured as UnmarshalErrorCollect,
		// this option limits the number of errors which are collected before
		// parsing is aborted.
		// Configure a limit of 0 to collect all errors, without upper limit.
		// Defaults to 10.
		MaxUnmarshalErrors uint64
		// Handler function for unmarshal errors during row parsing.
		// Takes precedence over all UnmarshalErrorHandling except
		// UnmarshalErrorIgnore.
		// Defaults to nil.
		RowUnmarshalErrorHandler RowUnmarshalErrorHandlerFunc
		// Handler function for columns not present in struct.
		// Defaults to nil.
		UnusedColumnsHandler UnusedColumnsHandlerFunc
	}
	UnmarshalErrorHandling uint8
	FieldError             struct {
		RowIndex     int // 0-based row index. Printed as 1-based row number in error text.
		ColumnIndex  int // 0-based column index.
		ColumnHeader string
		Err          error
	}
	ContentError struct {
		FieldErrors  []FieldError
		LimitReached bool
	}
)

var (
	// Ensure FieldError implements the error interface
	_ error = FieldError{}
	// Ensure FieldError can be unwrapped
	_ interface {
		Unwrap() error
	} = FieldError{}
	// Ensure ContentError implements the error interface
	_ error = ContentError{}
)

// Error implements error.
func (e FieldError) Error() string {
	return fmt.Sprintf("error unmarshalling column \"%s\" in row %d: %s", e.ColumnHeader, e.RowIndex+1, e.Err.Error())
}

// Unwrap
// Error implements the anonymous unwrap interface used by errors.Unwrap and others.
func (e FieldError) Unwrap() error {
	return e.Err
}

// Error implements error.
func (e ContentError) Error() string {
	if e.LimitReached {
		return fmt.Sprintf("too many (%d) errors reading data from Excel", len(e.FieldErrors))
	} else {
		return fmt.Sprintf("%d errors reading data from Excel", len(e.FieldErrors))
	}
}

// Unwrap
// Error implements the anonymous unwrap interface used by errors.Unwrap and others.
func (e ContentError) Unwrap() []error {
	// Slice needs to be type-adjusted
	errs := make([]error, len(e.FieldErrors))
	for i, v := range e.FieldErrors {
		errs[i] = v
	}
	return errs
}

const (
	// UnmarshalErrorIgnore
	// Ignore any errors during unmarshalling
	UnmarshalErrorIgnore UnmarshalErrorHandling = iota
	// UnmarshalErrorAbort
	// Abort reading when encountering the first unmarshalling error
	UnmarshalErrorAbort
	// UnmarshalErrorCollect
	// Collect unmarshalling errors up to a limit, but continue reading.
	// Collected errors are returned as one error at the end, of type
	UnmarshalErrorCollect
)

var (
	defaultReadConfig = func() *ReadConfig {
		return &ReadConfig{
			TagName:                "excel",
			DataStartRowIndex:      1,
			SkipUnknownColumns:     true,
			UnmarshalErrorHandling: UnmarshalErrorAbort,
			MaxUnmarshalErrors:     10,
		}
	}
	ErrSheetIndexOutOfRange        = errors.New("exl: sheet index out of range")
	ErrHeaderRowIndexOutOfRange    = errors.New("exl: header row index out of range")
	ErrDataStartRowIndexOutOfRange = errors.New("exl: data start row index out of range")
	ErrNoUnmarshaler               = errors.New("no unmarshaler")
	ErrNoDestinationField          = errors.New("no destination field with matching tag")
)

func readStrings(maxCol int, row *xlsx.Row) []string {
	ls := make([]string, maxCol)
	for i := 0; i < maxCol; i++ {
		ls[i] = row.GetCell(i).Value
	}
	return ls
}

func GetUnmarshalFunc(destField reflect.Value) UnmarshalExcelFunc {
	if destField.CanInterface() {

		inf := getFieldInterface(destField)
		if inf != nil {

			// Prefer ExcelUnmarshaler, if implemented
			if _, ok := inf.(ExcelUnmarshaler); ok {
				return UnmarshalExcelUnmarshaler
			}

			// Then handle specific types with special implementation
			if destField.Type() == reflect.TypeOf(time.Time{}) {
				return UnmarshalTime
			}

			// Then utilize TextUnmarshaler, e.g. for things like decimal.Decimal
			if _, ok := inf.(encoding.TextUnmarshaler); ok {
				return UnmarshalTextUnmarshaler
			}

		}
	}

	// And for primitive types, use custom unmarshalling func
	kind := destField.Type().Kind()
	isPointer := false
	if kind == reflect.Ptr {
		kind = destField.Type().Elem().Kind()
		isPointer = true
	}
	unmarshalFunc, ok := DefaultUnmarshalFuncs[kind]
	if ok {
		if isPointer {
			return func(destValue reflect.Value, cell *xlsx.Cell, params *ExcelUnmarshalParameters) error {
				reflect.New(destField.Type())
				return unmarshalPointer(destValue, cell, params, unmarshalFunc)
			}
		}
		return unmarshalFunc
	}

	return nil
}

func unmarshalPointer(destPointer reflect.Value, cell *xlsx.Cell, params *ExcelUnmarshalParameters, unmarshalFunc UnmarshalExcelFunc) error {
	// Create new pointer to the field value,
	// as the pointer may be nil
	elemType := destPointer.Type().Elem()
	destPointer.Set(reflect.New(elemType))

	// Unmarshal into that new value
	destValue := destPointer.Elem()
	return unmarshalFunc(destValue, cell, params)
}

// Read opens an xlsx file from the given io.Reader.
// Each row is parsed and unmarshalled into a slice of `T`.
// Note that this function needs to read the reader entirely
// into memory to determine the size, otherwise the zip reader cannot be called.
// Use one of the other `Read*` methods to avoid reading the whole file into memory
// before parsing starts - the excel library will copy the file content into memory anyways.
func Read[T ReadConfigurator](reader io.Reader, filterFunc ...func(t T) (add bool)) ([]T, error) {
	// since io.Reader does not provide a size, we have to read it all to get the size
	if bytes, err := io.ReadAll(reader); err != nil {
		return []T(nil), err
	} else {
		return ReadBinary(bytes, filterFunc...)
	}
}

// ReadReaderAt opens an xlsx file at the given file path.
// Each row is parsed and unmarshalled into a slice of `T`.
func ReadReaderAt[T ReadConfigurator](reader io.ReaderAt, size int64, filterFunc ...func(t T) (add bool)) ([]T, error) {
	f, err := xlsx.OpenReaderAt(reader, size)
	if err != nil {
		return nil, err
	}
	return ReadParsed[T](f, filterFunc...)
}

// ReadFile opens an xlsx file at the given file path.
// Each row is parsed and unmarshalled into a slice of `T`.
func ReadFile[T ReadConfigurator](file string, filterFunc ...func(t T) (add bool)) ([]T, error) {
	f, err := xlsx.OpenFile(file)
	if err != nil {
		return nil, err
	}
	return ReadParsed[T](f, filterFunc...)
}

// ReadBinary opens an xlsx file from the provided bytes.
// Each row is parsed and unmarshalled into a slice of `T`.
func ReadBinary[T ReadConfigurator](bytes []byte, filterFunc ...func(t T) (add bool)) ([]T, error) {
	f, err := xlsx.OpenBinary(bytes)
	if err != nil {
		return nil, err
	}
	return ReadParsed[T](f, filterFunc...)
}

type FieldInfo struct {
	reflectFieldIndex int
	Header            string
	unmarshalFunc     UnmarshalExcelFunc
}

// ReadParsed opens an already parsed xlsx file directly.
// Each row is parsed and unmarshalled into a slice of `T`.
func ReadParsed[T ReadConfigurator](f *xlsx.File, filterFunc ...func(t T) (add bool)) ([]T, error) {
	var t T
	rc := defaultReadConfig()
	t.ReadConfigure(rc)
	sidx := rc.SheetIndex
	if len(rc.SheetName) > 0 {
		for idx, s := range f.Sheets {
			if s.Name == rc.SheetName {
				sidx = idx
				break
			}
		}
	}
	if sidx < 0 || sidx > len(f.Sheet)-1 {
		return nil, ErrSheetIndexOutOfRange
	}
	sheet := f.Sheets[sidx]
	if rc.HeaderRowIndex < 0 || rc.HeaderRowIndex > sheet.MaxRow-1 {
		return nil, ErrHeaderRowIndexOutOfRange
	}
	if rc.DataStartRowIndex < 0 || rc.DataStartRowIndex > sheet.MaxRow-1 {
		return nil, ErrDataStartRowIndexOutOfRange
	}
	headerRow, _ := sheet.Row(rc.HeaderRowIndex)
	maxCol := sheet.MaxCol
	headers := readStrings(maxCol, headerRow)

	// Key: Header / Tag name
	// Value: Reflection field index
	tagToFieldMap := make(map[string]int)
	// Key: Column Index
	// Value: Unmarshalling Info
	columnFields := make([]FieldInfo, len(headers))

	typ := reflect.TypeOf(t).Elem()
	for i := 0; i < typ.NumField(); i++ {
		if ta := typ.Field(i).Tag; ta != "" {
			if tt, have := ta.Lookup(rc.TagName); have {
				tagToFieldMap[tt] = i
			}
		}
	}

	{
		val := reflect.New(typ).Elem()

		for columnIndex, header := range headers {
			reflectFieldIndex, have := tagToFieldMap[header]
			if !have {
				if rc.SkipUnknownColumns {
					// Skip reading this field
					columnFields[columnIndex] = FieldInfo{
						reflectFieldIndex: reflectFieldIndex,
						Header:            header,
						unmarshalFunc:     nil,
					}
					continue
				} else {
					return nil, fmt.Errorf("%w for column \"%s\" at index %d", ErrNoDestinationField, header, columnIndex)
				}
			}

			field := val.Field(reflectFieldIndex)

			unmarshaler := GetUnmarshalFunc(field)
			if unmarshaler == nil {
				if rc.SkipUnknownTypes {
					// Skip reading this field
					columnFields[columnIndex] = FieldInfo{
						reflectFieldIndex: reflectFieldIndex,
						Header:            header,
						unmarshalFunc:     nil,
					}
					continue
				} else {
					return nil, fmt.Errorf("%w for column \"%s\" at index %d", ErrNoUnmarshaler, header, columnIndex)
				}
			}

			columnFields[columnIndex] = FieldInfo{
				reflectFieldIndex: reflectFieldIndex,
				Header:            header,
				unmarshalFunc:     unmarshaler,
			}
		}
	}

	unmarshalConfig := &ExcelUnmarshalParameters{
		TrimSpace:           rc.TrimSpace,
		Date1904:            f.Date1904,
		FallbackDateFormats: rc.FallbackDateFormats,
	}

	collectedErrors := make([]FieldError, 0)

	ts := make([]T, 0)
	for rowIndex := 0; rowIndex < sheet.MaxRow; rowIndex++ {
		if rowIndex >= rc.DataStartRowIndex {
			val := reflect.New(typ).Elem()
			if row, _ := sheet.Row(rowIndex); row != nil {

				for columnIndex, fi := range columnFields {
					// If there is no unmarshal function,
					// this field has been skipped by previous logic.
					// e.g. no destination field, or unknown type.
					if fi.unmarshalFunc == nil {
						if rc.UnusedColumnsHandler != nil {
							rc.UnusedColumnsHandler(row.GetCell(columnIndex), &val, fi)
						}
						continue
					}
					cell := row.GetCell(columnIndex)

					destField := val.Field(fi.reflectFieldIndex)
					err := fi.unmarshalFunc(destField, cell, unmarshalConfig)
					if err != nil && rc.UnmarshalErrorHandling != UnmarshalErrorIgnore {
						if rc.RowUnmarshalErrorHandler != nil {
							rc.RowUnmarshalErrorHandler(cell, &val, fi)
							continue
						}
						fer := FieldError{
							RowIndex:     rowIndex,
							ColumnIndex:  columnIndex,
							ColumnHeader: fi.Header,
							Err:          err,
						}
						if rc.UnmarshalErrorHandling == UnmarshalErrorAbort {
							return nil, fer
						} else {
							collectedErrors = append(collectedErrors, fer)
							if rc.MaxUnmarshalErrors > 0 && uint64(len(collectedErrors)) >= rc.MaxUnmarshalErrors {
								return nil, ContentError{
									FieldErrors:  collectedErrors,
									LimitReached: true,
								}
							}
						}
					}
				}
				nT := val.Addr().Interface().(T)
				add := true
				if filterFunc != nil && len(filterFunc) > 0 {
					for _, fF := range filterFunc {
						if fF != nil {
							add = fF(nT)
							if !add {
								break
							}
						}
					}
				}
				if add {
					ts = append(ts, nT)
				}
			}
		}
	}
	if len(collectedErrors) > 0 {
		return nil, ContentError{
			FieldErrors:  collectedErrors,
			LimitReached: false,
		}
	}
	return ts, nil
}

// ReadExcel walk func from excel
func ReadExcel(file string, sheetIndex int, walk func(index int, rows *xlsx.Row)) error {
	f, err := xlsx.OpenFile(file)
	if err != nil {
		return err
	}
	sheet := f.Sheets[sheetIndex]
	for i := 0; i < sheet.MaxRow; i++ {
		if row, _ := sheet.Row(i); row != nil {
			walk(i, row)
		}
	}
	return nil
}

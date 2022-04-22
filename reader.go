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
	"errors"
	"io"
	"io/ioutil"
	"reflect"
	"strings"

	"github.com/billcoding/reflectx"
	"github.com/tealeg/xlsx/v3"
)

var (
	errMetadataIsNil               = errors.New("exl: the metadata is nil")
	errSheetIndexOutOfRange        = errors.New("exl: sheet index out of range")
	errHeaderRowIndexOutOfRange    = errors.New("exl: header row index out of range")
	errDataStartRowIndexOutOfRange = errors.New("exl: data start row index out of range")
)

func read(maxCol int, row *xlsx.Row) []string {
	ls := make([]string, maxCol, maxCol)
	for i := 0; i < maxCol; i++ {
		ls[i] = row.GetCell(i).Value
	}
	return ls
}

// Read defines read io.Reader each row bind to `T`
//
// params: file,excel file full path
//
// params: typed parameter T, must be implements exl.Bind
//
// params: filterFunc, filter callback func
func Read[T ReadBind](reader io.Reader, bind T, filterFunc ...func(t T) (add bool)) ([]T, error) {
	if bytes, err := io.ReadAll(reader); err != nil {
		return []T(nil), err
	} else {
		return ReadBinary(bytes, bind, filterFunc...)
	}
}

// ReadFile defines read excel each row bind to `T`
//
// params: file,excel file full path
//
// params: typed parameter T, must be implements exl.Bind
//
// params: filterFunc, filter callback func
func ReadFile[T ReadBind](file string, bind T, filterFunc ...func(t T) (add bool)) ([]T, error) {
	if bytes, err := ioutil.ReadFile(file); err != nil {
		return []T(nil), err
	} else {
		return ReadBinary(bytes, bind, filterFunc...)
	}
}

// ReadBinary defines read binary each row bind to `T`
//
// params: file,excel file full path
//
// params: typed parameter T, must be implements exl.Bind
//
// params: filterFunc, filter callback func
func ReadBinary[T ReadBind](bytes []byte, bind T, filterFunc ...func(t T) (add bool)) ([]T, error) {
	f, err := xlsx.OpenBinary(bytes)
	if err != nil {
		return nil, err
	}
	md := bind.ReadMetadata()
	if md == nil {
		return nil, errMetadataIsNil
	}
	if md.SheetIndex < 0 || md.SheetIndex > len(f.Sheet)-1 {
		return nil, errSheetIndexOutOfRange
	}
	sheet := f.Sheets[md.SheetIndex]
	if md.HeaderRowIndex < 0 || md.HeaderRowIndex > sheet.MaxRow-1 {
		return nil, errHeaderRowIndexOutOfRange
	}
	if md.DataStartRowIndex < 0 || md.DataStartRowIndex > sheet.MaxRow-1 {
		return nil, errDataStartRowIndexOutOfRange
	}
	trimSpace := md.TrimSpace
	headerRow, _ := sheet.Row(md.HeaderRowIndex)
	maxCol := sheet.MaxCol
	headers := read(maxCol, headerRow)
	headerMap := make(map[int]string, maxCol)
	for i, h := range headers {
		headerMap[i] = h
	}
	fieldMap := make(map[string]int, 0)
	typ := reflect.TypeOf(bind).Elem()
	tagName := "excel"
	if ta := md.TagName; ta != "" {
		tagName = ta
	}
	for i := 0; i < typ.NumField(); i++ {
		if t := typ.Field(i).Tag; t != "" {
			if tt, have := t.Lookup(tagName); have {
				fieldMap[tt] = i
			}
		}
	}
	ts := make([]T, 0)
	for i := 0; i < sheet.MaxRow; i++ {
		if i >= md.DataStartRowIndex {
			val := reflect.New(typ).Elem()
			if row, _ := sheet.Row(i); row != nil {
				for di, d := range read(maxCol, row) {
					if header, have := headerMap[di]; have {
						if fi, fa := fieldMap[header]; fa {
							fie := val.Field(fi)
							reflectx.SetValue(reflect.ValueOf(d), fie)
							if trimSpace && (fie.Type().Kind() == reflect.String ||
								(fie.Type().Kind() == reflect.Ptr && fie.Type().Elem().Kind() == reflect.String)) {
								fie.SetString(strings.TrimSpace(fie.String()))
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
	return ts, nil
}

// ReadExcel defines read walk func from excel
//
// params: file, excel file pull path
//
// params: sheetIndex, current sheet index
//
// params: walk, walk func
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

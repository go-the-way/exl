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
func Read[T ReadBind](reader io.Reader, bind T) ([]T, error) {
	if bytes, err := io.ReadAll(reader); err != nil {
		return []T(nil), err
	} else {
		return ReadBinary(bytes, bind)
	}
}

// ReadFile defines read excel each row bind to `T`
//
// params: file,excel file full path
//
// params: typed parameter T, must be implements exl.Bind
func ReadFile[T ReadBind](file string, bind T) ([]T, error) {
	if bytes, err := ioutil.ReadFile(file); err != nil {
		return []T(nil), err
	} else {
		return ReadBinary(bytes, bind)
	}
}

// ReadBinary defines read binary each row bind to `T`
//
// params: file,excel file full path
//
// params: typed parameter T, must be implements exl.Bind
func ReadBinary[T ReadBind](bytes []byte, bind T) ([]T, error) {
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
							reflectx.SetValue(reflect.ValueOf(d), val.Field(fi))
						}
					}
				}
				ts = append(ts, val.Addr().Interface().(T))
			}
		}
	}
	return ts, nil
}

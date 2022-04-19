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
	"github.com/tealeg/xlsx/v3"
	"reflect"
)

var (
	errTsIsNil = errors.New("exl: ts is nil")
)

func write(sheet *xlsx.Sheet, data []any) {
	r := sheet.AddRow()
	for _, cell := range data {
		r.AddCell().SetValue(cell)
	}
}

// Write defines write []T to excel file
//
// params: file,excel file full path
//
// params: typed parameter T, must be implements exl.Bind
func Write[T WriteBind](file string, ts []T) error {
	if ts == nil {
		return errTsIsNil
	}
	f := xlsx.NewFile()
	if len(ts) > 0 {
		md := ts[0].WriteMetadata()
		sheetName := "Sheet1"
		if md != nil && md.SheetName != "" {
			sheetName = md.SheetName
		}
		if sheet, _ := f.AddSheet(sheetName); sheet != nil {
			typ := reflect.TypeOf(ts[0]).Elem()
			numField := typ.NumField()
			header := make([]any, numField, numField)
			tagName := "excel"
			if md != nil && md.TagName != "" {
				tagName = md.TagName
			}
			for i := 0; i < numField; i++ {
				fe := typ.Field(i)
				name := fe.Name
				if tt, have := fe.Tag.Lookup(tagName); have {
					name = tt
				}
				header[i] = name
			}
			// write header
			write(sheet, header)
			// write data
			for _, t := range ts {
				data := make([]any, numField, numField)
				for i := 0; i < numField; i++ {
					data[i] = reflect.ValueOf(t).Elem().Field(i).Interface()
				}
				write(sheet, data)
			}
		}
	}

	if err := f.Save(file); err != nil {
		return err
	}

	return nil
}

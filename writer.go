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
	"github.com/tealeg/xlsx/v3"
	"os"
	"path/filepath"
	"reflect"
)

type File struct{ file *xlsx.File }

func NewFile(options ...xlsx.FileOption) *File {
	return &File{xlsx.NewFile(options...)}
}

func (f *File) Save(dist string) error {
	_ = os.MkdirAll(filepath.Dir(dist), 0600)
	return f.file.Save(dist)
}

func write(sheet *xlsx.Sheet, data []any) {
	r := sheet.AddRow()
	for _, cell := range data {
		r.AddCell().SetValue(cell)
	}
}

// Write defines write []T to Excel file
//
// params: file,Excel file full path
//
// params: typed parameter T, must be implements exl.Bind
func Write[T WriteBind](file string, ts []T) error {
	f := NewFile()
	WriteSheet[T](f, ts)
	return f.Save(file)
}

// WriteExcel defines write [][]string to excel
//
// params: dist, excel file pull path
//
// params: data, write data to excel
func WriteExcel(dist string, data [][]string) error {
	return WriteExcelSheets(dist, []SheetData{{"Sheet1", data}})
}

type SheetData struct {
	Sheet string
	Data  [][]string
}

func WriteExcelSheets(dist string, sheets []SheetData) error {
	f := NewFile()
	for _, sht := range sheets {
		if sheet, err := f.file.AddSheet(sht.Sheet); err != nil {
			return err
		} else {
			for _, row := range sht.Data {
				r := sheet.AddRow()
				for _, cell := range row {
					r.AddCell().SetString(cell)
				}
			}
		}
	}
	return f.Save(dist)
}

func WriteSheet[T WriteBind](file *File, ts []T) {
	wm := defaultWM()
	if len(ts) > 0 {
		ts[0].Write(wm)
	}
	tT := new(T)
	if sheet, _ := file.file.AddSheet(wm.SheetName); sheet != nil {
		typ := reflect.TypeOf(tT).Elem().Elem()
		numField := typ.NumField()
		header := make([]any, numField, numField)
		for i := 0; i < numField; i++ {
			fe := typ.Field(i)
			name := fe.Name
			if tt, have := fe.Tag.Lookup(wm.TagName); have {
				name = tt
			}
			header[i] = name
		}
		// write header
		write(sheet, header)
		if len(ts) > 0 {
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
}

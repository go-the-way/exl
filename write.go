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
	"io"
	"reflect"

	"github.com/tealeg/xlsx/v3"
)

type (
	WriteConfigurator interface{ WriteConfigure(wc *WriteConfig) }
	WriteConfig       struct {
		// Name of the Sheet created to hold the data.
		// Defaults to "Sheet1".
		SheetName string
		// Name of the tag on the data struct to configure column headers
		// and whether to ignore any fields.
		// Defaults to "excel".
		TagName string
		// If true, fields without the tag defined via TagName are ignored.
		// They are not written to the output file,
		// and will also not write a header.
		// Defaults to "false".
		IgnoreFieldsWithoutTag bool
	}
)

var defaultWriteConfig = func() *WriteConfig {
	return &WriteConfig{SheetName: "Sheet1", TagName: "excel", IgnoreFieldsWithoutTag: false}
}

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
func Write[T WriteConfigurator](file string, ts []T) error {
	f := xlsx.NewFile()
	write0(f, ts)
	return f.Save(file)
}

// WriteTo defines write to []T to excel file
//
// params: w, the dist writer
//
// params: typed parameter T, must be implements exl.Bind
func WriteTo[T WriteConfigurator](w io.Writer, ts []T) error {
	f := xlsx.NewFile()
	write0(f, ts)
	return f.Write(w)
}

func write0[T WriteConfigurator](f *xlsx.File, ts []T) {
	wc := defaultWriteConfig()
	tT := new(T)
	// Always configure writes, even if the provided data is empty.
	// If not done this way, empty files could have different headers
	// compared to files with content, because the write config would not run.
	(*tT).WriteConfigure(wc)
	if sheet, _ := f.AddSheet(wc.SheetName); sheet != nil {
		typ := reflect.TypeOf(tT).Elem().Elem()
		numField := typ.NumField()
		header := make([]any, 0, numField)
		ignoreField := make([]bool, numField)
		for i := 0; i < numField; i++ {
			fe := typ.Field(i)
			name := fe.Name
			if tt, have := fe.Tag.Lookup(wc.TagName); have {
				name = tt
			} else if wc.IgnoreFieldsWithoutTag {
				ignoreField[i] = true
				continue
			}
			header = append(header, name)
		}
		// write header
		write(sheet, header)
		if len(ts) > 0 {
			// write data
			for _, t := range ts {
				data := make([]any, 0, numField)
				for i := 0; i < numField; i++ {
					if !ignoreField[i] {
						data = append(data, reflect.ValueOf(t).Elem().Field(i).Interface())
					}
				}
				write(sheet, data)
			}
		}
	}
}

// WriteExcel defines write [][]string to excel
//
// params: file, excel file pull path
//
// params: data, write data to excel
func WriteExcel(file string, data [][]string) error {
	f := xlsx.NewFile()
	writeExcel0(f, data)
	return f.Save(file)
}

// WriteExcelTo defines write [][]string to excel
//
// params: w, the dist writer
//
// params: data, write data to excel
func WriteExcelTo(w io.Writer, data [][]string) error {
	f := xlsx.NewFile()
	writeExcel0(f, data)
	return f.Write(w)
}

func writeExcel0(f *xlsx.File, data [][]string) {
	sheet, _ := f.AddSheet("Sheet1")
	for _, row := range data {
		r := sheet.AddRow()
		for _, cell := range row {
			r.AddCell().SetString(cell)
		}
	}
}

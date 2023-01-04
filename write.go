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
	WriteConfigurator interface{ Configurator[*WriteConfig] }
	WriteConfig       struct {
		SheetName string
		TagName   string
	}
)

var defaultWriteConfig = func() *WriteConfig { return &WriteConfig{SheetName: "Sheet1", TagName: "excel"} }

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
	if len(ts) > 0 {
		ts[0].Configure(wc)
	}
	tT := new(T)
	if sheet, _ := f.AddSheet(wc.SheetName); sheet != nil {
		typ := reflect.TypeOf(tT).Elem().Elem()
		numField := typ.NumField()
		header := make([]any, numField, numField)
		for i := 0; i < numField; i++ {
			fe := typ.Field(i)
			name := fe.Name
			if tt, have := fe.Tag.Lookup(wc.TagName); have {
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

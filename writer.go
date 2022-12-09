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
	"fmt"
	"io"
	"reflect"

	"github.com/tealeg/xlsx/v3"
)

// Writer define a writer for exl
type Writer struct {
	file      *xlsx.File
	mapHeader []reflect.Value
	ignore    map[int]struct{}
}

// NewWriter returns new exl writer
func NewWriter(options ...xlsx.FileOption) *Writer {
	w := &Writer{file: xlsx.NewFile(options...)}
	w.reset()
	return w
}

// Write or append the param data into sheet
func (w *Writer) Write(sheet string, data any) error {
	if sht, ok := w.file.Sheet[sheet]; ok {
		w.reset()
		return w.writeSheet(sht, data)
	}
	if sht, err := w.file.AddSheet(sheet); err != nil {
		return err
	} else {
		w.reset()
		return w.writeSheet(sht, data)
	}
}

// SaveTo the buffered binary into dist file
func (w *Writer) SaveTo(path string) (err error) { return w.file.Save(path) }

// WriteTo the buffered binary into new writer
func (w *Writer) WriteTo(dw io.Writer) (n int, err error) { return 0, w.file.Write(dw) }

func (w *Writer) writeSheet(sheet *xlsx.Sheet, data any) (err error) {
	value := w.deepValue(reflect.ValueOf(data))
	vk := value.Type().Kind()
	switch vk {
	case reflect.Array, reflect.Slice:
		w.writeArrayOrSlice(sheet, value)
		return nil
	}
	return errors.New(fmt.Sprintf("not supported type: %v", vk))
}

func (w *Writer) writeArrayOrSlice(sheet *xlsx.Sheet, value reflect.Value) {
	arrLen := value.Len()
	var header *reflect.Value
	if arrLen > 0 {
		dv := w.deepValue(value.Index(0))
		header = &dv
	}
	w.setHeaderRow(sheet.AddRow(), w.deepType(value.Type().Elem()), header)
	for i := 0; i < arrLen; i++ {
		w.setDataRow(sheet.AddRow(), w.deepValue(value.Index(i)))
	}
}

func (w *Writer) setHeaderRow(row *xlsx.Row, typ reflect.Type, value *reflect.Value) {
	vk := typ.Kind()

	if vk == reflect.Map && value != nil {
		for _, _mk := range value.MapKeys() {
			mk := w.deepValue(_mk)
			w.mapHeader = append(w.mapHeader, mk)
			w.addCell(row, mk)
		}
		return
	}

	if vk == reflect.Struct {
		for i := 0; i < typ.NumField(); i++ {
			field := typ.Field(i)
			excelTag, _ := field.Tag.Lookup("excel")
			if excelTag == "" {
				row.AddCell().SetString(field.Name)
			} else if excelTag != "-" {
				row.AddCell().SetString(excelTag)
			}
			if excelTag == "-" {
				w.ignore[i] = struct{}{}
			}
		}
		return
	}

	row.AddCell().SetString("Unnamed")
}

func (w *Writer) setDataRow(row *xlsx.Row, value reflect.Value) {
	vk := value.Kind()

	if vk == reflect.Map {
		for _, k := range w.mapHeader {
			v := value.MapIndex(k)
			w.addCell(row, v)
		}
		return
	}

	if vk == reflect.Struct {
		da := value.Type()
		for i := 0; i < da.NumField(); i++ {
			if _, ok := w.ignore[i]; !ok {
				w.addCell(row, w.deepValue(value.Field(i)))
			}
		}
		return
	}

	w.addCell(row, value)
}

func (w *Writer) reset() {
	w.mapHeader = make([]reflect.Value, 0)
	w.ignore = make(map[int]struct{}, 0)
}

func (w *Writer) deepType(typ reflect.Type) reflect.Type {
	if typ.Kind() == reflect.Ptr {
		return w.deepType(typ.Elem())
	}
	return typ
}

func (w *Writer) deepValue(value reflect.Value) reflect.Value {
	if value.Type().Kind() == reflect.Ptr {
		return w.deepValue(value.Elem())
	}
	return value
}

func (w *Writer) addCell(row *xlsx.Row, value reflect.Value) {
	if value.CanInterface() {
		row.AddCell().SetValue(value.Interface())
	}
}

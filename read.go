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
	"os"
	"reflect"
	"strings"

	"github.com/tealeg/xlsx/v3"
)

type (
	Configurator[C any] interface {
		Configure(c C)
	}
	ReadConfigurator interface{ Configurator[*ReadConfig] }
	ReadConfig       struct {
		TagName           string
		SheetIndex        int
		HeaderRowIndex    int
		DataStartRowIndex int
		TrimSpace         bool
	}
)

var defaultReadConfig = func() *ReadConfig { return &ReadConfig{TagName: "excel", DataStartRowIndex: 1} }

var (
	ErrSheetIndexOutOfRange        = errors.New("exl: sheet index out of range")
	ErrHeaderRowIndexOutOfRange    = errors.New("exl: header row index out of range")
	ErrDataStartRowIndexOutOfRange = errors.New("exl: data start row index out of range")
)

func read(maxCol int, row *xlsx.Row) []string {
	ls := make([]string, maxCol, maxCol)
	for i := 0; i < maxCol; i++ {
		ls[i] = row.GetCell(i).Value
	}
	return ls
}

// Read io.Reader each row bind to `T`
func Read[T ReadConfigurator](reader io.Reader, filterFunc ...func(t T) (add bool)) ([]T, error) {
	if bytes, err := io.ReadAll(reader); err != nil {
		return []T(nil), err
	} else {
		return ReadBinary(bytes, filterFunc...)
	}
}

// ReadFile each row bind to `T`
func ReadFile[T ReadConfigurator](file string, filterFunc ...func(t T) (add bool)) ([]T, error) {
	if bytes, err := os.ReadFile(file); err != nil {
		return []T(nil), err
	} else {
		return ReadBinary(bytes, filterFunc...)
	}
}

// ReadBinary each row bind to `T`
func ReadBinary[T ReadConfigurator](bytes []byte, filterFunc ...func(t T) (add bool)) ([]T, error) {
	f, err := xlsx.OpenBinary(bytes)
	if err != nil {
		return nil, err
	}
	var t T
	rc := defaultReadConfig()
	t.Configure(rc)
	if rc.SheetIndex < 0 || rc.SheetIndex > len(f.Sheet)-1 {
		return nil, ErrSheetIndexOutOfRange
	}
	sheet := f.Sheets[rc.SheetIndex]
	if rc.HeaderRowIndex < 0 || rc.HeaderRowIndex > sheet.MaxRow-1 {
		return nil, ErrHeaderRowIndexOutOfRange
	}
	if rc.DataStartRowIndex < 0 || rc.DataStartRowIndex > sheet.MaxRow-1 {
		return nil, ErrDataStartRowIndexOutOfRange
	}
	trimSpace := rc.TrimSpace
	headerRow, _ := sheet.Row(rc.HeaderRowIndex)
	maxCol := sheet.MaxCol
	headers := read(maxCol, headerRow)
	headerMap := make(map[int]string, maxCol)
	for i, h := range headers {
		headerMap[i] = h
	}
	fieldMap := make(map[string]int, 0)
	typ := reflect.TypeOf(t).Elem()
	for i := 0; i < typ.NumField(); i++ {
		if ta := typ.Field(i).Tag; ta != "" {
			if tt, have := ta.Lookup(rc.TagName); have {
				fieldMap[tt] = i
			}
		}
	}
	ts := make([]T, 0)
	for i := 0; i < sheet.MaxRow; i++ {
		if i >= rc.DataStartRowIndex {
			val := reflect.New(typ).Elem()
			if row, _ := sheet.Row(i); row != nil {
				for di, d := range read(maxCol, row) {
					if header, have := headerMap[di]; have {
						if fi, fa := fieldMap[header]; fa {
							fie := val.Field(fi)
							setValue(reflect.ValueOf(d), fie)
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

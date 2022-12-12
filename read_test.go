// Copyright 2022 exl Author. All Rights Reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//	http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing percissions and
// limitations under the License.
package exl

import (
	"errors"
	"fmt"
	"os"
	"reflect"
	"strings"
	"testing"

	"github.com/tealeg/xlsx/v3"
)

type (
	readTmp struct {
		Name1 string `excel:"Name1"`
		Name2 string `excel:"Name2"`
		Name3 string `excel:"Name3"`
		Name4 string `excel:"Name4"`
		Name5 string `excel:"Name5"`
	}
	readSheetIndexOutOfRange        struct{}
	readHeaderRowIndexOutOfRange    struct{}
	readDataStartRowIndexOutOfRange struct{}
)

func (t *readTmp) Configure(rc *ReadConfig) {
	rc.TrimSpace = true
}

func (t *readSheetIndexOutOfRange) Configure(rc *ReadConfig) {
	rc.SheetIndex = -1
}

func (t *readHeaderRowIndexOutOfRange) Configure(rc *ReadConfig) {
	rc.HeaderRowIndex = -1
}

func (t *readDataStartRowIndexOutOfRange) Configure(rc *ReadConfig) {
	rc.DataStartRowIndex = -1
}

func TestReadFileErr(t *testing.T) {
	if _, err := ReadFile[*readTmp](""); err == nil {
		t.Error("test failed")
	}
	testFile := "tmp.xlsx"
	defer func() { _ = os.Remove(testFile) }()
	_ = Write(testFile, []*writeTmp{{}})
	if _, err := ReadFile[*readSheetIndexOutOfRange](testFile); err != ErrSheetIndexOutOfRange {
		t.Error("test failed")
	}
	if _, err := ReadFile[*readHeaderRowIndexOutOfRange](testFile); err != ErrHeaderRowIndexOutOfRange {
		t.Error("test failed")
	}
	if _, err := ReadFile[*readDataStartRowIndexOutOfRange](testFile); err != ErrDataStartRowIndexOutOfRange {
		t.Error("test failed")
	}
}

func TestReadBinaryErr(t *testing.T) {
	if _, err := ReadBinary[*readTmp](nil); err == nil {
		t.Error("test failed")
	}
}

type _reader struct{}

func (*_reader) Read([]byte) (n int, err error) {
	return 0, errors.New("")
}

func TestRead(t *testing.T) {
	if _, err := Read[*readTmp](&_reader{}); err == nil {
		t.Error("test failed")
	}
	if _, err := Read[*readTmp](strings.NewReader("")); err == nil {
		t.Error("test failed")
	}
}

func TestReadFile(t *testing.T) {
	testFile := "tmp.xlsx"
	defer func() { _ = os.Remove(testFile) }()
	data := [][]string{
		{"Name1", "Name2", "Name3", "Name4", "Name5"},
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
	}
	if err := WriteExcel(testFile, data); err != nil {
		t.Error("test failed: " + err.Error())
	}
	if models, err := ReadFile[*readTmp](testFile); err != nil {
		t.Error("test failed: " + err.Error())
	} else if len(models) != len(data)-1 {
		t.Error("test failed")
	} else {
		for i, m := range models {
			d := data[i+1]
			if d[0] != m.Name1 {
				t.Error("test failed: Name1 not equal")
			}
			if d[1] != m.Name2 {
				t.Error("test failed: Name2 not equal")
			}
			if d[2] != m.Name3 {
				t.Error("test failed: Name3 not equal")
			}
			if d[3] != m.Name4 {
				t.Error("test failed: Name4 not equal")
			}
			if d[4] != m.Name5 {
				t.Error("test failed: Name5 not equal")
			}
		}
	}
}

func TestReadTrimSpace(t *testing.T) {
	testFile := "tmp.xlsx"
	defer func() { _ = os.Remove(testFile) }()
	data := [][]string{
		{"Name1", "Name2", "Name3", "Name4", "Name5"},
		{"Name1 ", "Name2", "Name3", "Name4", "Name5"},
		{"Name11", "Name22 ", "Name33", "Name44", "Name55"},
		{"Name111", "Name222 ", "Name333 ", "Name444", "Name555"},
	}
	if err := WriteExcel(testFile, data); err != nil {
		t.Error("test failed: " + err.Error())
	}

	if models, err := ReadFile[*readTmp](testFile); err != nil {
		t.Error("test failed: " + err.Error())
	} else if models[0].Name1 != "Name1" || models[1].Name2 != "Name22" || models[2].Name3 != "Name333" {
		t.Error("test failed")
	}
}

func TestReadFilterFunc(t *testing.T) {
	testFile := "tmp.xlsx"
	defer func() { _ = os.Remove(testFile) }()
	data := [][]string{
		{"Name1", "Name2", "Name3", "Name4", "Name5"},
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
	}
	if err := WriteExcel(testFile, data); err != nil {
		t.Error("test failed: " + err.Error())
	}
	{
		if models, err := ReadFile[*readTmp](testFile, func(t *readTmp) (add bool) {
			return true
		}); err != nil {
			t.Error("test failed: " + err.Error())
		} else if len(models) != 2 {
			t.Error("test failed")
		}
	}
	{
		if models, err := ReadFile[*readTmp](testFile, func(t *readTmp) (add bool) {
			return false
		}); err != nil {
			t.Error("test failed: " + err.Error())
		} else if len(models) != 0 {
			t.Error("test failed")
		}
	}
	{
		if models, err := ReadFile[*readTmp](testFile, func(t *readTmp) (add bool) {
			return t.Name1 == "Name11"
		}); err != nil {
			t.Error("test failed: " + err.Error())
		} else if len(models) != 1 {
			t.Error("test failed")
		}
	}
}

func TestReadExcel(t *testing.T) {
	if err := ReadExcel("", 0, nil); err == nil {
		t.Error("test failed")
	}
}

func testBasic(testNum int) error {
	testFile := "tmp.xlsx"
	defer func() { _ = os.Remove(testFile) }()
	data := make([][]string, testNum, testNum)
	for i := range data {
		data[i] = []string{fmt.Sprintf("%d", i)}
	}
	if err := WriteExcel(testFile, data); err != nil {
		return err
	}
	if err := ReadExcel(testFile, 0, func(index int, rows *xlsx.Row) {
		if !reflect.DeepEqual(rows.GetCell(0).Value, fmt.Sprintf("%d", index)) {
			panic("test failed")
		}
	}); err != nil {
		return err
	}
	return nil
}

func TestBasic(t *testing.T) {
	_ = testBasic(10)
	_ = testBasic(100)
	_ = testBasic(10000)
}

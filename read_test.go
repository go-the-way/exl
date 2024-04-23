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
	"path"
	"reflect"
	"strings"
	"testing"
	"time"

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
	readErrorTmp struct {
		Name1       string `excel:"Name1"`
		Name2       int    `excel:"Name2"`
		ErrorsCount int
  }
	readUnusedTmp struct {
		Name1 string `excel:"Name1"`
		Name2 string `excel:"Name2"`
		Count int
	}
	readSheetIndexOutOfRange        struct{}
	readHeaderRowIndexOutOfRange    struct{}
	readDataStartRowIndexOutOfRange struct{}
)

func (t *readTmp) ReadConfigure(rc *ReadConfig) {
	rc.TrimSpace = true
}

func countUnmarshalErrors(cell *xlsx.Cell, val *reflect.Value, fi FieldInfo) {
	countF := val.FieldByName("ErrorsCount")
	countF.SetInt(countF.Int() + 1)
}

func (t *readErrorTmp) ReadConfigure(rc *ReadConfig) {
	rc.RowUnmarshalErrorHandler = countUnmarshalErrors
}

func countUnusedColumns(cell *xlsx.Cell, val *reflect.Value, fi FieldInfo) {
	countF := val.FieldByName("Count")
	countF.SetInt(countF.Int() + 1)
}

func (t *readUnusedTmp) ReadConfigure(rc *ReadConfig) {
	rc.UnusedColumnsHandler = countUnusedColumns
}

func (t *readSheetIndexOutOfRange) ReadConfigure(rc *ReadConfig) {
	rc.SheetIndex = -1
}

func (t *readHeaderRowIndexOutOfRange) ReadConfigure(rc *ReadConfig) {
	rc.HeaderRowIndex = -1
}

func (t *readDataStartRowIndexOutOfRange) ReadConfigure(rc *ReadConfig) {
	rc.DataStartRowIndex = -1
}

func TestFieldErrorError(t *testing.T) {
	fieldError := FieldError{
		RowIndex:     2,
		ColumnIndex:  7,
		ColumnHeader: "ColumnX",
		Err:          errors.New("unit test error"),
	}

	equal(t, "error unmarshalling column \"ColumnX\" in row 3: unit test error", fieldError.Error())
}

func TestFieldErrorIs(t *testing.T) {
	errUnit := errors.New("unit test error")
	fieldError := FieldError{
		Err: errUnit,
	}

	if !errors.Is(fieldError, errUnit) {
		t.Error("FieldError unwrapping failed")
	}
}

func TestFieldErrorUnwrap(t *testing.T) {
	errUnit := errors.New("unit test error")
	fieldError := FieldError{
		Err: errUnit,
	}

	unwrapped := fieldError.Unwrap()
	equal(t, errUnit, unwrapped)
}

func TestContentErrorError(t *testing.T) {
	t.Run("with limit reached", func(t *testing.T) {
		contentError := ContentError{
			FieldErrors: []FieldError{
				{}, {},
			},
			LimitReached: true,
		}
		equal(t, "too many (2) errors reading data from Excel", contentError.Error())
	})

	t.Run("without limit reached", func(t *testing.T) {
		contentError := ContentError{
			FieldErrors: []FieldError{
				{}, {},
			},
			LimitReached: false,
		}
		equal(t, "2 errors reading data from Excel", contentError.Error())
	})
}

func TestContentErrorUnwrap(t *testing.T) {
	errUnit1 := errors.New("unit test error 1")
	errUnit2 := errors.New("unit test error 2")
	contentError := ContentError{
		FieldErrors: []FieldError{
			{
				Err: errUnit1,
			},
			{
				Err: errUnit2,
			},
		},
	}

	expected := []error{
		FieldError{
			Err: errUnit1,
		},
		FieldError{
			Err: errUnit2,
		},
	}
	unwrapped := contentError.Unwrap()
	equal(t, expected, unwrapped)
}

type customUnmarshalledString string

func (s *customUnmarshalledString) UnmarshalExcel(cell *xlsx.Cell, params *ExcelUnmarshalParameters) error {
	if cell.Value == "error please" {
		return errors.New("excel unmarshalled: unit test error")
	} else {
		*s = customUnmarshalledString("excel unmarshalled: " + cell.Value)
		return nil
	}
}

type textUnmarshalledString string

func (s *textUnmarshalledString) UnmarshalText(text []byte) error {
	strValue := string(text)
	if strValue == "error please" {
		return errors.New("text unmarshalled: unit test error")
	} else {
		*s = textUnmarshalledString("text unmarshalled: " + strValue)
		return nil
	}
}

func TestGetUnmarshalFunc(t *testing.T) {
	type TestStruct struct {
		ExcelUnmarshalled            customUnmarshalledString
		TextUnmarshalled             textUnmarshalledString
		TimeUnmarshalled             time.Time
		PrimitiveUnmarshalled        string
		PrimitivePointerUnmarshalled *string
	}

	testStruct := &TestStruct{}
	val := reflect.ValueOf(testStruct).Elem()

	// Test cell with a value to be unmarshalled,
	// using a date value so the time unmarshaler can use this.
	// Every other unmarshaler will just use the raw string value
	successfulCell := &xlsx.Cell{
		Value:  "12000",
		NumFmt: xlsx.DefaultDateTimeFormat,
	}
	// Test cell with a specific value which causes the dummy unmarshalers
	// to explicitly cause errors, and the string unmarshaler to error out due to a formatting issue.
	errorCell := &xlsx.Cell{
		Value:  "error please",
		NumFmt: "<><><>error<><><>",
	}

	params := &ExcelUnmarshalParameters{}

	t.Run("ExcelUnmarshaler", func(t *testing.T) {
		field := val.FieldByName("ExcelUnmarshalled")
		unmarshaler := GetUnmarshalFunc(field)
		if unmarshaler == nil {
			t.Fatal("expected an unmarshaler func, got nil")
		}

		t.Run("successful", func(t *testing.T) {
			err := unmarshaler(field, successfulCell, params)
			if err != nil {
				t.Error("unexpected error:", err)
			}
			equal(t, customUnmarshalledString("excel unmarshalled: 12000"), testStruct.ExcelUnmarshalled)
		})
		t.Run("error", func(t *testing.T) {
			err := unmarshaler(field, errorCell, params)
			equal(t, "excel unmarshalled: unit test error", err.Error())
		})
	})

	t.Run("TextUnmarshaler", func(t *testing.T) {
		field := val.FieldByName("TextUnmarshalled")
		unmarshaler := GetUnmarshalFunc(field)
		if unmarshaler == nil {
			t.Fatal("expected an unmarshaler func, got nil")
		}

		t.Run("successful", func(t *testing.T) {
			err := unmarshaler(field, successfulCell, params)
			if err != nil {
				t.Error("unexpected error:", err)
			}
			equal(t, textUnmarshalledString("text unmarshalled: 12000"), testStruct.TextUnmarshalled)
		})
		t.Run("error", func(t *testing.T) {
			err := unmarshaler(field, errorCell, params)
			equal(t, "text unmarshalled: unit test error", err.Error())
		})
	})

	t.Run("Time", func(t *testing.T) {
		field := val.FieldByName("TimeUnmarshalled")
		unmarshaler := GetUnmarshalFunc(field)
		if unmarshaler == nil {
			t.Fatal("expected an unmarshaler func, got nil")
		}

		t.Run("successful", func(t *testing.T) {
			err := unmarshaler(field, successfulCell, params)
			if err != nil {
				t.Error("unexpected error:", err)
			}
			equal(t, time.Date(1932, time.November, 7, 0, 0, 0, 0, time.UTC), testStruct.TimeUnmarshalled)
		})
		t.Run("error", func(t *testing.T) {
			err := unmarshaler(field, errorCell, params)
			equal(t, "error parsing cell as date/time value: no recognized format", err.Error())
		})
	})
	t.Run("Primitive", func(t *testing.T) {
		field := val.FieldByName("PrimitiveUnmarshalled")
		unmarshaler := GetUnmarshalFunc(field)
		if unmarshaler == nil {
			t.Fatal("expected an unmarshaler func, got nil")
		}

		t.Run("successful", func(t *testing.T) {
			err := unmarshaler(field, successfulCell, params)
			if err != nil {
				t.Error("unexpected error:", err)
			}
			equal(t, "12000", testStruct.PrimitiveUnmarshalled)
		})
		t.Run("error", func(t *testing.T) {
			err := unmarshaler(field, errorCell, params)
			equal(t, "error formatting string value: invalid formatting code: unsupported or unescaped characters", err.Error())
		})
	})
	t.Run("Primitive Pointer", func(t *testing.T) {
		field := val.FieldByName("PrimitivePointerUnmarshalled")
		unmarshaler := GetUnmarshalFunc(field)
		if unmarshaler == nil {
			t.Fatal("expected an unmarshaler func, got nil")
		}

		t.Run("successful", func(t *testing.T) {
			err := unmarshaler(field, successfulCell, params)
			if err != nil {
				t.Error("unexpected error:", err)
			}
			expected := "12000"
			equal(t, &expected, testStruct.PrimitivePointerUnmarshalled)
		})
		t.Run("error", func(t *testing.T) {
			err := unmarshaler(field, errorCell, params)
			equal(t, "error formatting string value: invalid formatting code: unsupported or unescaped characters", err.Error())
		})
	})
}

type _reader struct{}

func (*_reader) Read([]byte) (n int, err error) {
	return 0, errors.New("read: unit test error")
}

func (*_reader) ReadAt([]byte, int64) (n int, err error) {
	return 0, errors.New("readat: unit test error")
}

func TestReadBinary(t *testing.T) {
	t.Run("error reading empty slice", func(t *testing.T) {
		_, err := ReadBinary[*readTmp]([]byte{})
		equal(t, "zip: not a valid zip file", err.Error())
	})
	t.Run("error (not panic) reading nil", func(t *testing.T) {
		_, err := ReadBinary[*readTmp](nil)
		equal(t, "zip: not a valid zip file", err.Error())
	})
	t.Run("success reading valid file", func(t *testing.T) {
		testFile := path.Join(t.TempDir(), "tmp.xlsx")

		// Shortened test data, see tests for ReadFile for full test

		data := [][]string{
			{"Name1", "Name2", "Name3", "Name4", "Name5"},
			{"Name11", "Name22", "Name33", "Name44", "Name55"},
		}

		if err := WriteExcel(testFile, data); err != nil {
			t.Error("test failed: " + err.Error())
		}

		bytes, err := os.ReadFile(testFile)
		if err != nil {
			t.Error("test failed: " + err.Error())
		}

		if models, err := ReadBinary[*readTmp](bytes); err != nil {
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
	})
}

func TestRead(t *testing.T) {
	t.Run("error reading empty reader", func(t *testing.T) {
		_, err := Read[*readTmp](strings.NewReader(""))
		equal(t, "zip: not a valid zip file", err.Error())
	})
	t.Run("error reading with read error", func(t *testing.T) {
		_, err := Read[*readTmp](&_reader{})
		equal(t, "read: unit test error", err.Error())
	})
	t.Run("success reading valid file", func(t *testing.T) {
		testFile := path.Join(t.TempDir(), "tmp.xlsx")

		// Shortened test data, see tests for ReadFile for full test

		data := [][]string{
			{"Name1", "Name2", "Name3", "Name4", "Name5"},
			{"Name11", "Name22", "Name33", "Name44", "Name55"},
		}

		if err := WriteExcel(testFile, data); err != nil {
			t.Error("test failed: " + err.Error())
		}

		reader, err := os.Open(testFile)
		if err != nil {
			t.Error("test failed: " + err.Error())
		}
		defer reader.Close()

		if models, err := Read[*readTmp](reader); err != nil {
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
	})
}

func TestReadReaderAt(t *testing.T) {
	t.Run("error reading empty reader", func(t *testing.T) {
		_, err := ReadReaderAt[*readTmp](strings.NewReader(""), 0)
		equal(t, "zip: not a valid zip file", err.Error())
	})
	t.Run("error reading with read error", func(t *testing.T) {
		_, err := ReadReaderAt[*readTmp](&_reader{}, 0)
		equal(t, "readat: unit test error", err.Error())
	})
	t.Run("success reading valid file", func(t *testing.T) {
		testFile := path.Join(t.TempDir(), "tmp.xlsx")

		// Shortened test data, see tests for ReadFile for full test

		data := [][]string{
			{"Name1", "Name2", "Name3", "Name4", "Name5"},
			{"Name11", "Name22", "Name33", "Name44", "Name55"},
		}

		if err := WriteExcel(testFile, data); err != nil {
			t.Error("test failed: " + err.Error())
		}

		reader, err := os.Open(testFile)
		if err != nil {
			t.Error("test failed: " + err.Error())
		}
		defer reader.Close()
		fi, err := os.Stat(testFile)
		if err != nil {
			t.Error("test failed: " + err.Error())
		}

		if models, err := ReadReaderAt[*readTmp](reader, fi.Size()); err != nil {
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
	})
}

func TestReadFile(t *testing.T) {

	t.Run("error reading missing file", func(t *testing.T) {
		testFile := path.Join(t.TempDir(), "does_not_exist.xlsx")
		_, err := ReadFile[*readTmp](testFile)
		// Concrete error message may vary between systems,
		// only verifying first part from xlsx library
		equal(t, "OpenFile", strings.Split(err.Error(), ":")[0])
	})

	t.Run("error reading with sheet index out of range", func(t *testing.T) {
		testFile := path.Join(t.TempDir(), "tmp.xlsx")
		err := Write(testFile, []*writeTmp{{}})
		if err != nil {
			t.Fatal(err.Error())
		}

		_, err = ReadFile[*readSheetIndexOutOfRange](testFile)
		equal(t, ErrSheetIndexOutOfRange, err)
	})

	t.Run("error reading with header row index out of range", func(t *testing.T) {
		testFile := path.Join(t.TempDir(), "tmp.xlsx")
		err := Write(testFile, []*writeTmp{{}})
		if err != nil {
			t.Fatal(err.Error())
		}

		_, err = ReadFile[*readHeaderRowIndexOutOfRange](testFile)
		equal(t, ErrHeaderRowIndexOutOfRange, err)
	})

	t.Run("error reading with start row index out of range", func(t *testing.T) {
		testFile := path.Join(t.TempDir(), "tmp.xlsx")
		err := Write(testFile, []*writeTmp{{}})
		if err != nil {
			t.Fatal(err.Error())
		}

		_, err = ReadFile[*readDataStartRowIndexOutOfRange](testFile)
		equal(t, ErrDataStartRowIndexOutOfRange, err)
	})

	t.Run("successful reading", func(t *testing.T) {
		testFile := path.Join(t.TempDir(), "tmp.xlsx")

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
	})

}

func TestReadTrimSpace(t *testing.T) {
	testFile := path.Join(t.TempDir(), "tmp.xlsx")

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

func TestHandleUnusedColumns(t *testing.T) {
	testFile := "tmp.xlsx"
	defer func() { _ = os.Remove(testFile) }()
	data := [][]string{
		{"Name1", "Name2", "Name3", "Name4", "Name5"},
		{"Name1 ", "Name2", "Name3", "Name4", "Name5"},
	}
	if err := WriteExcel(testFile, data); err != nil {
		t.Error("test failed: " + err.Error())
	}

	models, err := ReadFile[*readUnusedTmp](testFile)
	if err != nil {
		t.Error("test failed: " + err.Error())
	}
	if len(models) != 1 || models[0].Count != 3 {
		t.Error("test failed")
	}
}

func TestHandleUnmarshalErrors(t *testing.T) {
	testFile := "tmp.xlsx"
	defer func() { _ = os.Remove(testFile) }()
	data := [][]string{
		{"Name1", "Name2"},
		{"Name1 ", "Something"},
		{"Name1 ", "22"},
	}
	if err := WriteExcel(testFile, data); err != nil {
		t.Error("test failed: " + err.Error())
	}

	models, err := ReadFile[*readErrorTmp](testFile)
	if err != nil {
		t.Error("test failed: " + err.Error())
	}
	if len(models) != 2 || models[0].ErrorsCount != 1 || models[1].ErrorsCount != 0 {
		t.Error("test failed")
	}
}

type missingColumnsAllowed struct {
	Name1 string `excel:"Name1"`
}

func (*missingColumnsAllowed) ReadConfigure(rc *ReadConfig) {
	rc.SkipUnknownColumns = true
}

type missingColumnsNotAllowed struct {
	Name1 string `excel:"Name1"`
}

func (*missingColumnsNotAllowed) ReadConfigure(rc *ReadConfig) {
	rc.SkipUnknownColumns = false
}

func TestReadSkipColumns(t *testing.T) {
	testFile := path.Join(t.TempDir(), "tmp.xlsx")

	data := [][]string{
		{"Name1", "Name2", "Name3", "Name4", "Name5"},
		{"Name1 ", "Name2", "Name3", "Name4", "Name5"},
		{"Name11", "Name22 ", "Name33", "Name44", "Name55"},
		{"Name111", "Name222 ", "Name333 ", "Name444", "Name555"},
	}
	if err := WriteExcel(testFile, data); err != nil {
		t.Error("test failed: " + err.Error())
	}

	t.Run("allow missing columns", func(t *testing.T) {
		if _, err := ReadFile[*missingColumnsAllowed](testFile); err != nil {
			t.Error("test failed:", err)
		}
	})
	t.Run("disallow missing columns", func(t *testing.T) {
		_, err := ReadFile[*missingColumnsNotAllowed](testFile)
		if err == nil {
			t.Error("test failed: expected error, got nil")
		} else {
			equal(t, "no destination field with matching tag for column \"Name2\" at index 1", err.Error())
		}
	})
}

type missingTypesAllowed struct {
	// Using error as field type,
	// as error is an interface and thus
	// not unmarshalable without concrete type
	Name1 error `excel:"Name1"`
}

func (*missingTypesAllowed) ReadConfigure(rc *ReadConfig) {
	rc.SkipUnknownTypes = true
}

type missingTypesNotAllowed struct {
	// Using error as field type,
	// as error is an interface and thus
	// not unmarshalable without concrete type
	Name1 error `excel:"Name1"`
}

func (*missingTypesNotAllowed) ReadConfigure(rc *ReadConfig) {
	rc.SkipUnknownTypes = false
}

func TestReadSkipTypes(t *testing.T) {
	testFile := path.Join(t.TempDir(), "tmp.xlsx")

	data := [][]string{
		{"Name1"},
		{"Name1 Content"},
	}
	if err := WriteExcel(testFile, data); err != nil {
		t.Error("test failed: " + err.Error())
	}

	t.Run("allow missing unmarshalers", func(t *testing.T) {
		if _, err := ReadFile[*missingTypesAllowed](testFile); err != nil {
			t.Error("test failed:", err)
		}
	})
	t.Run("disallow missing unmarshalers", func(t *testing.T) {
		_, err := ReadFile[*missingTypesNotAllowed](testFile)
		if err == nil {
			t.Error("test failed: expected error, got nil")
		} else {
			equal(t, "no unmarshaler for column \"Name1\" at index 0", err.Error())
		}
	})
}

type ignoreUnmarshalErrors struct {
	Name1 customUnmarshalledString `excel:"Name1"`
}

func (*ignoreUnmarshalErrors) ReadConfigure(rc *ReadConfig) {
	rc.UnmarshalErrorHandling = UnmarshalErrorIgnore
}

type abortUnmarshalErrors struct {
	Name1 customUnmarshalledString `excel:"Name1"`
}

func (*abortUnmarshalErrors) ReadConfigure(rc *ReadConfig) {
	rc.UnmarshalErrorHandling = UnmarshalErrorAbort
}

type collectUnmarshalErrors struct {
	Name1 customUnmarshalledString `excel:"Name1"`
}

func (*collectUnmarshalErrors) ReadConfigure(rc *ReadConfig) {
	rc.UnmarshalErrorHandling = UnmarshalErrorCollect
	rc.MaxUnmarshalErrors = 2
}

type collectUnmarshalErrorsUnlimited struct {
	Name1 customUnmarshalledString `excel:"Name1"`
}

func (*collectUnmarshalErrorsUnlimited) ReadConfigure(rc *ReadConfig) {
	rc.UnmarshalErrorHandling = UnmarshalErrorCollect
	rc.MaxUnmarshalErrors = 0
}

func TestUnmarshalErrors(t *testing.T) {
	testFile := path.Join(t.TempDir(), "tmp.xlsx")

	data := [][]string{
		{"Name1"},
		{"error please"},
		{"error please"},
		{"error please"},
	}
	if err := WriteExcel(testFile, data); err != nil {
		t.Error("test failed: " + err.Error())
	}

	t.Run("ignore unmarshal errors", func(t *testing.T) {
		model, err := ReadFile[*ignoreUnmarshalErrors](testFile)
		if err != nil {
			t.Error("test failed:", err)
		}
		equal(t, customUnmarshalledString(""), model[0].Name1)
	})
	t.Run("abort at first error", func(t *testing.T) {
		model, err := ReadFile[*abortUnmarshalErrors](testFile)
		if err == nil {
			t.Error("test failed: expected error, got nil")
		} else {
			equal(t, FieldError{
				RowIndex:     1,
				ColumnIndex:  0,
				ColumnHeader: "Name1",
				Err:          errors.New("excel unmarshalled: unit test error"),
			}, err)
			if model != nil {
				t.Error("test failed: expected nil result, got:", model)
			}
		}
	})
	t.Run("collect errors limited", func(t *testing.T) {
		model, err := ReadFile[*collectUnmarshalErrors](testFile)
		if err == nil {
			t.Error("test failed: expected error, got nil")
		} else {
			equal(t, ContentError{
				FieldErrors: []FieldError{
					{
						RowIndex:     1,
						ColumnIndex:  0,
						ColumnHeader: "Name1",
						Err:          errors.New("excel unmarshalled: unit test error"),
					},
					{
						RowIndex:     2,
						ColumnIndex:  0,
						ColumnHeader: "Name1",
						Err:          errors.New("excel unmarshalled: unit test error"),
					},
				},
				LimitReached: true,
			}, err)
			if model != nil {
				t.Error("test failed: expected nil result, got:", model)
			}
		}
	})
	t.Run("collect errors unlimited", func(t *testing.T) {
		model, err := ReadFile[*collectUnmarshalErrorsUnlimited](testFile)
		if err == nil {
			t.Error("test failed: expected error, got nil")
		} else {
			equal(t, ContentError{
				FieldErrors: []FieldError{
					{
						RowIndex:     1,
						ColumnIndex:  0,
						ColumnHeader: "Name1",
						Err:          errors.New("excel unmarshalled: unit test error"),
					},
					{
						RowIndex:     2,
						ColumnIndex:  0,
						ColumnHeader: "Name1",
						Err:          errors.New("excel unmarshalled: unit test error"),
					},
					{
						RowIndex:     3,
						ColumnIndex:  0,
						ColumnHeader: "Name1",
						Err:          errors.New("excel unmarshalled: unit test error"),
					},
				},
				LimitReached: false,
			}, err)
			if model != nil {
				t.Error("test failed: expected nil result, got:", model)
			}
		}
	})
}

func TestReadFilterFunc(t *testing.T) {
	testFile := path.Join(t.TempDir(), "tmp.xlsx")

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

func testBasic(t *testing.T, testNum int) {
	testFile := path.Join(t.TempDir(), "tmp.xlsx")

	data := make([][]string, testNum)
	for i := range data {
		data[i] = []string{fmt.Sprintf("%d", i)}
	}
	if err := WriteExcel(testFile, data); err != nil {
		t.Fatal(err.Error())
	}
	if err := ReadExcel(testFile, 0, func(index int, rows *xlsx.Row) {
		equal(t, fmt.Sprintf("%d", index), rows.GetCell(0).Value)
	}); err != nil {
		t.Fatal(err.Error())
	}
}

func TestBasic(t *testing.T) {
	testBasic(t, 10)
	testBasic(t, 100)
	testBasic(t, 10000)
}

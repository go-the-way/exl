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
	"math"
	"reflect"
	"testing"
	"time"

	"codeberg.org/tealeg/xlsx/v4"
)

type _model struct {
	S   string
	B   bool
	I   int64
	I8  int8
	U   uint64
	U8  uint8
	F   float64
	F32 float32
	T   time.Time
	EU  customUnmarshalledString
	TU  textUnmarshalledString

	A any
}

func TestUnmarshalString(t *testing.T) {
	model := &_model{}
	destField := reflect.ValueOf(model).Elem().FieldByName("S")
	cell := &xlsx.Cell{}

	t.Run("string cell", func(t *testing.T) {
		cell.SetValue("string value")
		err := UnmarshalString(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, "string value", model.S)
	})

	t.Run("formatted cell", func(t *testing.T) {
		cell.SetFloatWithFormat(17.3, "0.00e+00")
		err := UnmarshalString(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, "1.730000e+01", model.S)
	})

	t.Run("don't trim space if not configured", func(t *testing.T) {
		cell.SetValue("  string value  ")
		err := UnmarshalString(destField, cell, &ExcelUnmarshalParameters{
			TrimSpace: false,
		})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, "  string value  ", model.S)
	})

	t.Run("trim space if configured", func(t *testing.T) {
		cell.SetValue("  string value  ")
		err := UnmarshalString(destField, cell, &ExcelUnmarshalParameters{
			TrimSpace: true,
		})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, "string value", model.S)
	})
}

func TestUnmarshalBool(t *testing.T) {
	model := &_model{}
	destField := reflect.ValueOf(model).Elem().FieldByName("B")
	cell := &xlsx.Cell{}

	t.Run("true cell", func(t *testing.T) {
		cell.SetBool(true)
		err := UnmarshalBool(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, true, model.B)
	})

	t.Run("false cell", func(t *testing.T) {
		cell.SetBool(false)
		err := UnmarshalBool(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, false, model.B)
	})
}

func TestUnmarshalInt(t *testing.T) {
	model := &_model{}
	destField := reflect.ValueOf(model).Elem().FieldByName("I")
	cell := &xlsx.Cell{}

	t.Run("positive integer cell", func(t *testing.T) {
		cell.SetValue(123)
		err := UnmarshalInt(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, int64(123), model.I)
	})

	t.Run("negative integer cell", func(t *testing.T) {
		cell.SetValue(-123)
		err := UnmarshalInt(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, int64(-123), model.I)
	})

	t.Run("text cell", func(t *testing.T) {
		cell.SetValue("123")
		err := UnmarshalInt(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, int64(123), model.I)
	})

	t.Run("float cell", func(t *testing.T) {
		cell.SetValue(123.7)
		err := UnmarshalInt(destField, cell, &ExcelUnmarshalParameters{})
		if err == nil {
			t.Fatal("expected format error")
		}
	})

	t.Run("overflow", func(t *testing.T) {
		destFieldOverflow := reflect.ValueOf(model).Elem().FieldByName("I8")
		cell.SetValue(math.MaxInt8 + 1)
		err := UnmarshalInt(destFieldOverflow, cell, &ExcelUnmarshalParameters{})
		if err == nil {
			t.Fatal("expected overflow error")
		}
	})
}

func TestUnmarshalUInt(t *testing.T) {
	model := &_model{}
	destField := reflect.ValueOf(model).Elem().FieldByName("U")
	cell := &xlsx.Cell{}

	t.Run("positive integer cell", func(t *testing.T) {
		cell.SetValue(123)
		err := UnmarshalUInt(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, uint64(123), model.U)
	})

	t.Run("negative integer cell: error", func(t *testing.T) {
		cell.SetValue(-123)
		err := UnmarshalUInt(destField, cell, &ExcelUnmarshalParameters{})
		if err == nil {
			t.Fatal("expected error")
		}
	})

	t.Run("text cell", func(t *testing.T) {
		cell.SetValue("123")
		err := UnmarshalUInt(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, uint64(123), model.U)
	})

	t.Run("float cell", func(t *testing.T) {
		cell.SetValue(123.7)
		err := UnmarshalUInt(destField, cell, &ExcelUnmarshalParameters{})
		if err == nil {
			t.Fatal("expected format error")
		}
	})

	t.Run("overflow", func(t *testing.T) {
		destFieldOverflow := reflect.ValueOf(model).Elem().FieldByName("U8")
		cell.SetValue(math.MaxUint8 + 1)
		err := UnmarshalUInt(destFieldOverflow, cell, &ExcelUnmarshalParameters{})
		if err == nil {
			t.Fatal("expected overflow error")
		}
	})
}

func TestUnmarshalFloat(t *testing.T) {
	model := &_model{}
	destField := reflect.ValueOf(model).Elem().FieldByName("F")
	cell := &xlsx.Cell{}

	t.Run("positive float cell", func(t *testing.T) {
		cell.SetValue(123.7)
		err := UnmarshalFloat(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, 123.7, model.F)
	})

	t.Run("negative float cell", func(t *testing.T) {
		cell.SetValue(-123.7)
		err := UnmarshalFloat(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, -123.7, model.F)
	})

	t.Run("text cell with float value", func(t *testing.T) {
		cell.SetValue("123.7")
		err := UnmarshalFloat(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, 123.7, model.F)
	})

	t.Run("int cell", func(t *testing.T) {
		cell.SetValue(123)
		err := UnmarshalFloat(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, 123.0, model.F)
	})

	t.Run("int cell", func(t *testing.T) {
		destFieldOverflow := reflect.ValueOf(model).Elem().FieldByName("F32")
		cell.SetValue(math.MaxFloat64)
		err := UnmarshalFloat(destFieldOverflow, cell, &ExcelUnmarshalParameters{})
		if err == nil {
			t.Fatal("expected overflow error")
		}
	})

	t.Run("text cell with non-numeric value returns error", func(t *testing.T) {
		cell.SetValue("not a number")
		err := UnmarshalFloat(destField, cell, &ExcelUnmarshalParameters{})
		if err == nil {
			t.Fatal("expected error, got nil")
		}
		equal(t, "error parsing cell as float value: strconv.ParseFloat: parsing \"not a number\": invalid syntax", err.Error())
	})
}

func TestUnmarshalTime(t *testing.T) {
	model := &_model{}
	destField := reflect.ValueOf(model).Elem().FieldByName("T")
	cell := &xlsx.Cell{}

	// NOTE: Testing with too accurate of a date may result in floating point errors.
	testTime := time.Date(2023, time.November, 13, 14, 15, 0, 0, time.UTC)
	testTimeFormatted := testTime.Format(time.RFC3339)

	t.Run("date cell", func(t *testing.T) {
		cell.SetDate(testTime)
		err := UnmarshalTime(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, testTime, model.T)
	})

	t.Run("date cell with fallback value", func(t *testing.T) {
		cell.SetDate(testTime)
		cell.Value = testTimeFormatted
		err := UnmarshalTime(destField, cell, &ExcelUnmarshalParameters{
			FallbackDateFormats: []string{time.RFC3339},
		})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, testTime, model.T)
	})

	t.Run("date cell with unparsable value", func(t *testing.T) {
		cell.SetDate(testTime)
		cell.Value = "rubbish"
		err := UnmarshalTime(destField, cell, &ExcelUnmarshalParameters{})
		if err == nil {
			t.Fatal("expected error, got nil")
		}
	})

	t.Run("text cell with fallback value", func(t *testing.T) {
		cell.SetString(testTimeFormatted)
		err := UnmarshalTime(destField, cell, &ExcelUnmarshalParameters{
			FallbackDateFormats: []string{time.RFC3339},
		})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, testTime, model.T)
	})

	t.Run("text cell with unparsable value", func(t *testing.T) {
		cell.SetString("rubbish")
		err := UnmarshalTime(destField, cell, &ExcelUnmarshalParameters{})
		if err == nil {
			t.Fatal("expected error, got nil")
		}
	})
}

func TestUnmarshalExcelUnmarshaler(t *testing.T) {
	model := &_model{}
	destField := reflect.ValueOf(model).Elem().FieldByName("EU")
	cell := &xlsx.Cell{}

	t.Run("successful unmarshalling", func(t *testing.T) {
		cell.SetString("unit test value")
		err := UnmarshalExcelUnmarshaler(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, customUnmarshalledString("excel unmarshalled: unit test value"), model.EU)
	})

	t.Run("unsuccessful unmarshalling returns error", func(t *testing.T) {
		cell.SetString("error please") // Trigger an error in the unmarshaller
		err := UnmarshalExcelUnmarshaler(destField, cell, &ExcelUnmarshalParameters{})
		equal(t, errors.New("excel unmarshalled: unit test error"), err)
	})

	t.Run("catch wrong type", func(t *testing.T) {
		wrongField := reflect.ValueOf(model).Elem().FieldByName("TU")
		err := UnmarshalExcelUnmarshaler(wrongField, cell, &ExcelUnmarshalParameters{})
		equal(t, ErrCannotCastUnmarshaler, err)
	})
}

func TestUnmarshalTextUnmarshaler(t *testing.T) {
	model := &_model{}
	destField := reflect.ValueOf(model).Elem().FieldByName("TU")
	cell := &xlsx.Cell{}

	t.Run("successful unmarshalling", func(t *testing.T) {
		cell.SetString("unit test value")
		err := UnmarshalTextUnmarshaler(destField, cell, &ExcelUnmarshalParameters{})
		if err != nil {
			t.Fatal(err)
		}
		equal(t, textUnmarshalledString("text unmarshalled: unit test value"), model.TU)
	})

	t.Run("unsuccessful unmarshalling returns error", func(t *testing.T) {
		cell.SetString("error please") // Trigger an error in the unmarshaller
		err := UnmarshalTextUnmarshaler(destField, cell, &ExcelUnmarshalParameters{})
		equal(t, errors.New("text unmarshalled: unit test error"), err)
	})

	t.Run("catch wrong type", func(t *testing.T) {
		wrongField := reflect.ValueOf(model).Elem().FieldByName("EU")
		err := UnmarshalTextUnmarshaler(wrongField, cell, &ExcelUnmarshalParameters{})
		equal(t, ErrCannotCastUnmarshaler, err)
	})
}

func equal(t *testing.T, expected, actual any) {
	t.Helper()
	if !reflect.DeepEqual(expected, actual) {
		t.Errorf("test failed, expected \"%v\", got \"%v\"", expected, actual)
	}
}

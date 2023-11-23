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
	"encoding"
	"errors"
	"fmt"
	"reflect"
	"strings"
	"time"

	"github.com/tealeg/xlsx/v3"
)

var ErrNegativeUInt = errors.New("negative integer provided for unsigned field")
var ErrOverflow = errors.New("numeric overflow, number is too large for this field")
var ErrNoRecognizedFormat = errors.New("no recognized format")

// ErrCannotCastUnmarshaler is returned in case a field technically implements an unmarshaler interface,
// but casting to it at runtime failed for some reason.
var ErrCannotCastUnmarshaler = errors.New("cannot cast to unmarshaler interface")

var DefaultUnmarshalFuncs = map[reflect.Kind]UnmarshalExcelFunc{
	reflect.String:  UnmarshalString,
	reflect.Bool:    UnmarshalBool,
	reflect.Int:     UnmarshalInt,
	reflect.Int8:    UnmarshalInt,
	reflect.Int16:   UnmarshalInt,
	reflect.Int32:   UnmarshalInt,
	reflect.Int64:   UnmarshalInt,
	reflect.Uint:    UnmarshalUInt,
	reflect.Uintptr: UnmarshalUInt,
	reflect.Uint8:   UnmarshalUInt,
	reflect.Uint16:  UnmarshalUInt,
	reflect.Uint32:  UnmarshalUInt,
	reflect.Uint64:  UnmarshalUInt,
	reflect.Float32: UnmarshalFloat,
	reflect.Float64: UnmarshalFloat,
}

type ExcelUnmarshalParameters struct {
	// See ReadConfig.TrimSpace
	TrimSpace bool
	// See xlsx.File.Date1904
	Date1904 bool
	// See ReadConfig.FallbackDateFormats
	FallbackDateFormats []string
}

type ExcelUnmarshaler interface {
	UnmarshalExcel(cell *xlsx.Cell, params *ExcelUnmarshalParameters) error
}

type UnmarshalExcelFunc func(destValue reflect.Value, cell *xlsx.Cell, params *ExcelUnmarshalParameters) error

func UnmarshalString(destValue reflect.Value, cell *xlsx.Cell, params *ExcelUnmarshalParameters) error {
	str, err := cell.FormattedValue()
	if err != nil {
		return fmt.Errorf("error formatting string value: %w", err)
	}
	if params.TrimSpace {
		str = strings.TrimSpace(str)
	}
	destValue.SetString(str)
	return nil
}

func UnmarshalBool(destValue reflect.Value, cell *xlsx.Cell, params *ExcelUnmarshalParameters) error {
	destValue.SetBool(cell.Bool())
	return nil
}

func UnmarshalInt(destValue reflect.Value, cell *xlsx.Cell, params *ExcelUnmarshalParameters) error {
	val, err := cell.Int64()
	if err != nil {
		return fmt.Errorf("error parsing cell as integer value: %w", err)
	}
	if destValue.OverflowInt(val) {
		return ErrOverflow
	}
	destValue.SetInt(val)
	return nil
}

func UnmarshalUInt(destValue reflect.Value, cell *xlsx.Cell, params *ExcelUnmarshalParameters) error {
	val, err := cell.Int64()
	if err != nil {
		return fmt.Errorf("error parsing cell as integer value: %w", err)
	}
	if val < 0 {
		return ErrNegativeUInt
	}
	uval := uint64(val)
	if destValue.OverflowUint(uval) {
		return ErrOverflow
	}
	destValue.SetUint(uval)
	return nil
}

func UnmarshalFloat(destValue reflect.Value, cell *xlsx.Cell, params *ExcelUnmarshalParameters) error {
	val, err := cell.Float()
	if err != nil {
		return fmt.Errorf("error parsing cell as float value: %w", err)
	}
	if destValue.OverflowFloat(val) {
		return ErrOverflow
	}
	destValue.SetFloat(val)
	return nil
}

func UnmarshalTime(destValue reflect.Value, cell *xlsx.Cell, params *ExcelUnmarshalParameters) error {
	var val time.Time
	if cell.IsTime() {
		var err error
		val, err = cell.GetTime(params.Date1904)
		if err != nil {
			var ok bool
			val, ok = unmarshalTimeFallback(cell.Value, params.FallbackDateFormats)
			if !ok {
				return fmt.Errorf("error parsing cell as date/time value: %w", err)
			}
		}
	} else {
		var ok bool
		val, ok = unmarshalTimeFallback(cell.Value, params.FallbackDateFormats)
		if !ok {
			return fmt.Errorf("error parsing cell as date/time value: %w", ErrNoRecognizedFormat)
		}
	}
	destValue.Set(reflect.ValueOf(val))
	return nil
}

func unmarshalTimeFallback(value string, formats []string) (time.Time, bool) {
	for _, format := range formats {
		val, err := time.Parse(format, value)
		if err == nil {
			return val, true
		}
	}
	return time.Time{}, false
}

func getFieldInterface(destField reflect.Value) any {
	destFieldPointer := destField

	// Same logic as json.Unmarshal is using:
	// If the field has a named type and is addressable,
	// start with its address, so that if the type has pointer methods,
	// we find them.
	// Usually unmarshaler implementations are on the pointer type,
	// so that they can actually write back to the field when called.
	if destField.Kind() != reflect.Pointer && destField.Type().Name() != "" && destField.CanAddr() {
		destFieldPointer = destField.Addr()
	}

	return destFieldPointer.Interface()
}

func UnmarshalExcelUnmarshaler(destField reflect.Value, cell *xlsx.Cell, params *ExcelUnmarshalParameters) error {
	unmarshaler, ok := getFieldInterface(destField).(ExcelUnmarshaler)
	if !ok {
		// This should not happen at runtime,
		// as we have already cast successfully to get here
		return ErrCannotCastUnmarshaler
	}

	return unmarshaler.UnmarshalExcel(cell, params)
}

func UnmarshalTextUnmarshaler(destField reflect.Value, cell *xlsx.Cell, params *ExcelUnmarshalParameters) error {
	unmarshaler, ok := getFieldInterface(destField).(encoding.TextUnmarshaler)
	if !ok {
		// This should not happen at runtime,
		// as we have alread cast successfully to get here
		return ErrCannotCastUnmarshaler
	}

	return unmarshaler.UnmarshalText([]byte(cell.Value))
}

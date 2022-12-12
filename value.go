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
	"fmt"
	"reflect"
	"strconv"
)

func setValue(srcValue, distValue reflect.Value) {
	switch distValue.Kind() {
	case reflect.String:
		setStringValue(srcValue, distValue)
	case reflect.Bool:
		setBoolValue(srcValue, distValue)
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		setIntValue(srcValue, distValue)
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		setUintValue(srcValue, distValue)
	case reflect.Float32, reflect.Float64:
		setFloatValue(srcValue, distValue)
	default:
		if srcValue.Type() == distValue.Type() {
			distValue.Set(srcValue)
		}
	}
}

func setStringValue(srcValue, distValue reflect.Value) {
	switch srcValue.Kind() {
	// Interface->String
	default:
		if srcValue.CanInterface() {
			distValue.SetString(fmt.Sprintf("%v", srcValue.Interface()))
		}
	// String->String
	case reflect.String:
		distValue.SetString(srcValue.String())
	}
}

func setBoolValue(srcValue, distValue reflect.Value) {
	switch srcValue.Kind() {
	// String->Bool
	case reflect.String:
		boolVal, _ := strconv.ParseBool(srcValue.String())
		distValue.SetBool(boolVal)
	// Bool->Bool
	case reflect.Bool:
		distValue.Set(srcValue)
	}
}

// set int value
func setIntValue(srcValue, distValue reflect.Value) {
	switch srcValue.Kind() {
	// String->Int
	case reflect.String:
		floatVal, _ := strconv.ParseFloat(srcValue.String(), 64)
		distValue.SetInt(int64(floatVal))
	// Float->Int
	case reflect.Float32, reflect.Float64:
		distValue.SetInt(int64(srcValue.Float()))
	// Uint->Int
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		distValue.SetInt(int64(srcValue.Uint()))
	// Int->Int
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		distValue.SetInt(srcValue.Int())
	}
}

// set uint value
func setUintValue(srcValue, distValue reflect.Value) {
	switch srcValue.Kind() {
	// String->Uint
	case reflect.String:
		floatVal, _ := strconv.ParseFloat(srcValue.String(), 64)
		distValue.SetUint(uint64(floatVal))
	// Float->Uint
	case reflect.Float32, reflect.Float64:
		distValue.SetUint(uint64(srcValue.Float()))
	// Int->Uint
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		distValue.SetUint(uint64(srcValue.Int()))
	// Uint->Uint
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		distValue.SetUint(srcValue.Uint())
	}
}

// set float value
func setFloatValue(srcValue, distValue reflect.Value) {
	switch srcValue.Kind() {
	// String->Float
	case reflect.String:
		floatVal, _ := strconv.ParseFloat(srcValue.String(), 64)
		distValue.SetFloat(floatVal)
	// Int->Float
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		distValue.SetFloat(float64(srcValue.Int()))
	// Uint->Float
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		distValue.SetFloat(float64(srcValue.Uint()))
	// Float->Float
	case reflect.Float32, reflect.Float64:
		distValue.SetFloat(srcValue.Float())
	}
}

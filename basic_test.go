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
	"os"
	"reflect"
	"testing"

	"github.com/tealeg/xlsx/v3"
)

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

func BenchmarkBasic(b *testing.B) {
	for i := 0; i < b.N; i++ {
		_ = testBasic(10)
		_ = testBasic(100)
		_ = testBasic(10000)
	}
}

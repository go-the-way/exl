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
	"bytes"
	"os"
	"testing"
)

func TestWriter(t *testing.T) {
	defer os.Remove("out.xlsx")
	w := NewWriter()
	type testCase struct {
		name string
		data any
	}
	var ptr = 10
	var tcs = []testCase{
		{"notSupport", 1000},
		{"int", []int{1, 2}},
		{"int\\ / ? * [ ]", []int{1, 2}},
		{"int", []int{11, 22}},
		{"float", []float32{1.1, 2.2}},
		{"string", []string{"coco", "yoyo"}},
		{"boolean", []bool{true, false}},
		{"map", []map[string]any{{"key": "kk", "val": "vv"}}},
		{"ptr", []*map[string]any{{"key": "kkk", "val": "vvv"}}},
		{"struct", []struct {
			ID    int    `excel:"编号"`
			Name  string `excel:"名称"`
			Extra bool   `excel:"-"`
			Age   int    `excel:"年龄"`
			Addr  string
			IDPtr *int
		}{
			{10, "Apple", false, 25, "Addr1", &ptr},
			{20, "Pear", true, 26, "Addr2", &ptr},
			{30, "Banana", true, 30, "Addr3", &ptr},
		}},
	}
	for _, tc := range tcs {
		t.Run(tc.name, func(t *testing.T) {
			w.Write(tc.name, tc.data)
		})
	}
	_ = w.SaveTo("out.xlsx")
	_, _ = w.WriteTo(&bytes.Buffer{})
}

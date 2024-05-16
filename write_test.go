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
	"os"
	"path"
	"strings"
	"testing"
)

type writeTmp struct {
	Name1 string `excel:"Name1"`
	Name2 string `excel:"Name2"`
	Name3 string `excel:"Name3"`
	Name4 string `excel:"Name4"`
	Name5 string `excel:"Name5"`
}

type writeReadTmp writeTmp

func (*writeReadTmp) WriteConfigure(_ *WriteConfig) {}
func (*writeReadTmp) ReadConfigure(_ *ReadConfig)   {}
func (*writeTmp) WriteConfigure(_ *WriteConfig)     {}

// Row type which ignored fields without tag.
type writeWithIgnore struct {
	Name1 string // This is ignored, should not be written
	Name2 string `excel:"Name2"`
}

func (*writeWithIgnore) WriteConfigure(wc *WriteConfig) {
	wc.IgnoreFieldsWithoutTag = true
}

// Row type for negative test, when a field does not have a tag
// but the write configuration is set to write it anyway.
type writeWithoutIgnore struct {
	Name1 string // This is NOT ignored, should be written
	Name2 string `excel:"Name2"`
}

func (*writeWithoutIgnore) WriteConfigure(wc *WriteConfig) {
	wc.IgnoreFieldsWithoutTag = false
}

// Row type used below to inspect headers of otherwise empty files.
// The error message is checked to make sure ignored columns don't get a header.
type readNoFields struct {
}

func (*readNoFields) ReadConfigure(rc *ReadConfig) {
	rc.SkipUnknownColumns = false
	// "Read" the header to not get an error message immediately
	rc.DataStartRowIndex = 0
}

func TestWriteErr(t *testing.T) {
	testFile := "tmp.xlsx"
	defer func() { _ = os.Remove(testFile) }()
	if err := Write(testFile, []*writeTmp{}); err != nil {
		t.Error("test failed")
	}
}

func TestWrite(t *testing.T) {
	testFile := "tmp.xlsx"
	defer func() { _ = os.Remove(testFile) }()
	data := []*writeTmp{
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
	}
	if err := Write(testFile, data); err != nil {
		t.Error("test failed: " + err.Error())
	}
	if models, err := ReadFile[*writeReadTmp](testFile); err != nil {
		t.Error("test failed: " + err.Error())
	} else if len(models) != len(data) {
		t.Error("test failed")
	} else {
		for i, m := range models {
			d := data[i]
			if d.Name1 != m.Name1 {
				t.Error("test failed: Name1 not equal")
			}
			if d.Name2 != m.Name2 {
				t.Error("test failed: Name2 not equal")
			}
			if d.Name3 != m.Name3 {
				t.Error("test failed: Name3 not equal")
			}
			if d.Name4 != m.Name4 {
				t.Error("test failed: Name4 not equal")
			}
			if d.Name5 != m.Name5 {
				t.Error("test failed: Name5 not equal")
			}
		}
	}
}

func TestWriteTo(t *testing.T) {
	data := []*writeTmp{
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
	}
	testFile := "tmp.xlsx"
	file, err := os.Create(testFile)
	defer func() { _ = file.Close(); _ = os.Remove(testFile) }()
	if err != nil {
		t.Error("test failed: " + err.Error())
	}
	if err = WriteTo(file, data); err != nil {
		t.Error("test failed: " + err.Error())
	}
	if models, err := ReadFile[*writeReadTmp](testFile); err != nil {
		t.Error("test failed: " + err.Error())
	} else if len(models) != len(data) {
		t.Error("test failed")
	} else {
		for i, m := range models {
			d := data[i]
			if d.Name1 != m.Name1 {
				t.Error("test failed: Name1 not equal")
			}
			if d.Name2 != m.Name2 {
				t.Error("test failed: Name2 not equal")
			}
			if d.Name3 != m.Name3 {
				t.Error("test failed: Name3 not equal")
			}
			if d.Name4 != m.Name4 {
				t.Error("test failed: Name4 not equal")
			}
			if d.Name5 != m.Name5 {
				t.Error("test failed: Name5 not equal")
			}
		}
	}
}

func TestWriteExcelTo(t *testing.T) {
	data := []*writeTmp{
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
	}
	testFile := "tmp.xlsx"
	file, err := os.Create(testFile)
	defer func() { _ = file.Close(); _ = os.Remove(testFile) }()
	if err != nil {
		t.Error("test failed: " + err.Error())
	}
	if err = WriteExcelTo(file, [][]string{
		{"Name1", "Name2", "Name3", "Name4", "Name5"},
		{"Name11", "Name22", "Name33", "Name44", "Name55"},
		{"Name111", "Name222", "Name333", "Name444", "Name555"},
		{"Name1111", "Name2222", "Name3333", "Name4444", "Name5555"},
		{"Name11111", "Name22222", "Name33333", "Name44444", "Name55555"},
		{"Name111111", "Name222222", "Name333333", "Name444444", "Name555555"},
	}); err != nil {
		t.Error("test failed: " + err.Error())
	}
	if models, err := ReadFile[*writeReadTmp](testFile); err != nil {
		t.Error("test failed: " + err.Error())
	} else if len(models) != len(data) {
		t.Error("test failed")
	} else {
		for i, m := range models {
			d := data[i]
			if d.Name1 != m.Name1 {
				t.Error("test failed: Name1 not equal")
			}
			if d.Name2 != m.Name2 {
				t.Error("test failed: Name2 not equal")
			}
			if d.Name3 != m.Name3 {
				t.Error("test failed: Name3 not equal")
			}
			if d.Name4 != m.Name4 {
				t.Error("test failed: Name4 not equal")
			}
			if d.Name5 != m.Name5 {
				t.Error("test failed: Name5 not equal")
			}
		}
	}
}

func TestWriteIgnoringFieldsWithoutTag(t *testing.T) {
	testFile := path.Join(t.TempDir(), "tmp.xlsx")

	data := []*writeWithIgnore{
		{"Name11", "Name22"},
		{"Name111", "Name222"},
		{"Name1111", "Name2222"},
		{"Name11111", "Name22222"},
		{"Name111111", "Name222222"},
	}
	if err := Write(testFile, data); err != nil {
		t.Error("test failed: " + err.Error())
	}
	if models, err := ReadFile[*writeReadTmp](testFile); err != nil {
		t.Error("test failed: " + err.Error())
	} else if len(models) != len(data) {
		t.Error("test failed")
	} else {
		for i, m := range models {
			d := data[i]
			// Name2 should not be in the excel file, so also not read
			if m.Name1 != "" {
				t.Error("test failed: Name1 not empty")
			}
			if d.Name2 != m.Name2 {
				t.Error("test failed: Name2 not equal")
			}
		}
	}
}

func TestWriteNotIgnoringFieldsWithoutTag(t *testing.T) {
	testFile := path.Join(t.TempDir(), "tmp.xlsx")

	data := []*writeWithoutIgnore{
		{"Name11", "Name22"},
		{"Name111", "Name222"},
		{"Name1111", "Name2222"},
		{"Name11111", "Name22222"},
		{"Name111111", "Name222222"},
	}
	if err := Write(testFile, data); err != nil {
		t.Error("test failed: " + err.Error())
	}
	if models, err := ReadFile[*writeReadTmp](testFile); err != nil {
		t.Error("test failed: " + err.Error())
	} else if len(models) != len(data) {
		t.Error("test failed")
	} else {
		for i, m := range models {
			d := data[i]
			// Should not be ignored, even though it does not have a tag
			if d.Name1 != m.Name1 {
				t.Error("test failed: Name1 not equal")
			}
			if d.Name2 != m.Name2 {
				t.Error("test failed: Name2 not equal")
			}
		}
	}
}

// Ensure that for empty data sets, the write configuration is still done,
// so that custom tag names (if configured) are not ignored,
// and ignored fields (if configured) don't get headers.
func TestWriteWithoutRowsConfiguresHeader(t *testing.T) {
	testFile := path.Join(t.TempDir(), "tmp.xlsx")

	// Write data without content, which should write only a header
	data := []*writeWithIgnore{}
	if err := Write(testFile, data); err != nil {
		t.Error("test failed: " + err.Error())
	}

	// Read the file without any fields in the struct,
	// which will give us an error message for the first column,
	// which should be "Name2" as "Name1" is ignored and should not be in the file.
	_, err := ReadFile[*readNoFields](testFile)
	if err == nil {
		t.Error("test failed, expected error")
	}
	if !strings.Contains(err.Error(), "Name2") {
		t.Error("test failed, expected message for second column, got: " + err.Error())
	}
}

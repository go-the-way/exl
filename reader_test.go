package exl

import (
	"errors"
	"os"
	"strings"
	"testing"
)

type (
	readTmp struct {
		Name1 string `excel:"Name1"`
		Name2 string `excel:"Name2"`
		Name3 string `excel:"Name3"`
		Name4 string `excel:"Name4"`
		Name5 string `excel:"Name5"`
	}
	readNilMetadata                 struct{}
	readSheetIndexOutOfRange        struct{}
	readHeaderRowIndexOutOfRange    struct{}
	readDataStartRowIndexOutOfRange struct{}
)

func (*readTmp) ReadMetadata() *ReadMetadata {
	return &ReadMetadata{DataStartRowIndex: 1, TagName: "excel"}
}

func (*readNilMetadata) ReadMetadata() *ReadMetadata {
	return nil
}

func (*readSheetIndexOutOfRange) ReadMetadata() *ReadMetadata {
	return &ReadMetadata{SheetIndex: -1}
}

func (*readHeaderRowIndexOutOfRange) ReadMetadata() *ReadMetadata {
	return &ReadMetadata{HeaderRowIndex: -1}
}

func (*readDataStartRowIndexOutOfRange) ReadMetadata() *ReadMetadata {
	return &ReadMetadata{DataStartRowIndex: -1}
}

func TestReadFileErr(t *testing.T) {
	if _, err := ReadFile("", new(readTmp)); err == nil {
		t.Error("test failed")
	}
	testFile := "tmp.xlsx"
	defer func() { _ = os.Remove(testFile) }()
	_ = Write(testFile, []*writeTmp{{}})
	if _, err := ReadFile(testFile, new(readNilMetadata)); err != errMetadataIsNil {
		t.Error("test failed")
	}
	if _, err := ReadFile(testFile, new(readSheetIndexOutOfRange)); err != errSheetIndexOutOfRange {
		t.Error("test failed")
	}
	if _, err := ReadFile(testFile, new(readHeaderRowIndexOutOfRange)); err != errHeaderRowIndexOutOfRange {
		t.Error("test failed")
	}
	if _, err := ReadFile(testFile, new(readDataStartRowIndexOutOfRange)); err != errDataStartRowIndexOutOfRange {
		t.Error("test failed")
	}
}

func TestReadBinaryErr(t *testing.T) {
	if _, err := ReadBinary(nil, new(readTmp)); err == nil {
		t.Error("test failed")
	}
}

type _reader struct{}

func (*_reader) Read([]byte) (n int, err error) {
	return 0, errors.New("")
}

func TestRead(t *testing.T) {
	if _, err := Read(&_reader{}, new(readTmp)); err == nil {
		t.Error("test failed")
	}
	if _, err := Read(strings.NewReader(""), new(readTmp)); err == nil {
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
	if models, err := ReadFile(testFile, new(readTmp)); err != nil {
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

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

type (
	// ReadBind defines read bind metadata
	ReadBind interface{ Read(rm *ReadMetadata) }
	// WriteBind defines write bind metadata
	WriteBind interface{ Write(wm *WriteMetadata) }
	// ReadMetadata defines read metadata
	ReadMetadata struct {
		TagName           string // TagName: tag name
		SheetIndex        int    // SheetIndex: read sheet index
		HeaderRowIndex    int    // HeaderRowIndex: sheet header row index
		DataStartRowIndex int    // DataStartRowIndex: sheet data start row index
		TrimSpace         bool   // TrimSpace: trim space left and right only on `string` type
	}
	// WriteMetadata defines write metadata
	WriteMetadata struct {
		SheetName string // SheetName: default sheet name
		TagName   string // TagName: tag name
	}
)

var (
	defaultRM = func() *ReadMetadata { return &ReadMetadata{TagName: "excel", DataStartRowIndex: 1} }
	defaultWM = func() *WriteMetadata { return &WriteMetadata{SheetName: "Sheet1", TagName: "excel"} }
)

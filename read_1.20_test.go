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

//go:build go1.20

package exl

import (
	"errors"
	"testing"
)

// TestContentErrorIs tests unwrapping of errors with potentially more than one wrapped error.
// This is only supported starting go 1.20 (when errors.Join() was added).
func TestContentErrorIs(t *testing.T) {
	errUnit := errors.New("unit test error")
	contentError := ContentError{
		FieldErrors: []FieldError{
			{
				Err: errUnit,
			},
		},
	}

	if !errors.Is(contentError, errUnit) {
		t.Error("ContentError unwrapping failed")
	}
}

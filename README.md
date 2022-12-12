# exl
Excel binding to struct written in Go.(Only supports Go1.18+)

[![CircleCI](https://circleci.com/gh/go-the-way/exl/tree/main.svg?style=shield)](https://circleci.com/gh/go-the-way/exl/tree/main)
![GitHub go.mod Go version](https://img.shields.io/github/go-mod/go-version/go-the-way/exl)
[![codecov](https://codecov.io/gh/go-the-way/exl/branch/main/graph/badge.svg?token=8MAR3J959H)](https://codecov.io/gh/go-the-way/exl)
[![Go Report Card](https://goreportcard.com/badge/github.com/go-the-way/exl)](https://goreportcard.com/report/github.com/go-the-way/exl)
[![GoDoc](https://pkg.go.dev/badge/github.com/go-the-way/exl?status.svg)](https://pkg.go.dev/github.com/go-the-way/exl?tab=doc)
[![Mentioned in Awesome Go](https://awesome.re/mentioned-badge.svg)](https://github.com/avelino/awesome-go#microsoft-excel)

## usage

### Read Excel

```go
package main

import (
	"fmt"
	"github.com/go-the-way/exl"
)

type ReadExcel struct {
	ID   int    `excel:"ID"`
	Name string `excel:"Name"`
}

func (*ReadExcel) Configure(rc *exl.ReadConfig) {}

func main() {
	if models, err := exl.ReadFile[*ReadExcel]("/to/path.xlsx"); err != nil {
		fmt.Println("read excel err:" + err.Error())
	} else {
		fmt.Printf("read excel num: %d\n", len(models))
	}
}
```

### Write Excel

```go
package main

import (
	"fmt"
	
	"github.com/go-the-way/exl"
)

type WriteExcel struct {
	ID   int    `excel:"ID"`
	Name string `excel:"Name"`
}

func (m *WriteExcel) Configure(wc *exl.WriteConfig) {}

func main() {
	if err := exl.Write("/to/path.xlsx", []*WriteExcel{{100, "apple"}, {200, "pear"}}); err != nil {
		fmt.Println("write excel err:" + err.Error())
	} else {
		fmt.Println("write excel done")
	}
}
```

## Writer

```go
package main

import (
	"fmt"

	"github.com/go-the-way/exl"
)

func main() {
	w := exl.NewWriter()
	if err := w.Write("int", []int{1, 2}); err != nil {
		fmt.Println(err)
		return
	}
	if err := w.Write("float", []float64{1.1, 2.2}); err != nil {
		fmt.Println(err)
		return
	}
	if err := w.Write("string", []string{"hello", "world"}); err != nil {
		fmt.Println(err)
		return
	}
	if err := w.Write("map", []map[string]string{{"id":"1000","name":"hello"},{"id":"2000","name":"world"}}); err != nil {
		fmt.Println(err)
		return
	}
	if err := w.Write("structWithField", []struct{ID int}{{1000},{2000}}); err != nil {
		fmt.Println(err)
		return
	}
	if err := w.Write("structWithTag", []struct{ID int `excel:"编号"`}{{1000},{2000}}); err != nil {
		fmt.Println(err)
		return
	}
	if err := w.Write("structWithTagAndIgnore", []struct{
		ID int `excel:"编号"`
		Extra int `excel:"-"`
		Name string `excel:"名称"`
	}{{1000,0,"Coco"},{2000,0,"Apple"}}); err != nil {
		fmt.Println(err)
		return
	} 
	if err := w.SaveTo("dist.xlsx"); err != nil {
		fmt.Println(err)
		return
	}
}
```

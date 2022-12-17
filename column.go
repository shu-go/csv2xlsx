package main

import (
	"strings"
)

type columns []column

type column struct {
	Sheet string
	Name  string

	Type        colType
	InputFormat string
}

func newColumn(s string, typ colType, format string) column {
	c := column{Sheet: "", Name: s, Type: typ, InputFormat: format}
	if pos := strings.Index(s, "!"); pos != -1 {
		c.Sheet = s[:pos]
		c.Name = s[pos+1:]
	}
	return c
}

func (cc columns) findByName(sheet, name string) int {
	if cc == nil {
		return -1
	}

	for i := range cc {
		if strings.EqualFold(cc[i].Sheet, sheet) && strings.EqualFold(cc[i].Name, name) {
			return i
		}
	}

	for i := range cc {
		if strings.EqualFold(cc[i].Sheet, "") && strings.EqualFold(cc[i].Name, name) {
			return i
		}
	}

	return -1
}

type colType string

const (
	typeUnknown  colType = ""
	typeText     colType = "text"
	typeNumber   colType = "number"
	typeDate     colType = "date"
	typeTime     colType = "time"
	typeDatetime colType = "datetime"
)

package main

import (
	"path/filepath"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type columns []column

type column struct {
	Sheet string
	Name  string

	Type        colType
	InputFormat string
}

func newColumn(s string, typ colType, format string) column {
	s = strings.ToLower(s)
	c := column{Sheet: "", Name: s, Type: typ, InputFormat: format}
	if pos := strings.Index(s, "!"); pos != -1 {
		c.Sheet = s[:pos]
		c.Name = s[pos+1:]
	}
	return c
}

func (cc columns) findByName(sheet, name string, optColNum ...int) int {
	if cc == nil {
		return -1
	}

	colNameAlpha := ""
	colNameNum := ""
	if len(optColNum) > 0 {
		colNameNum = "#" + strconv.Itoa(optColNum[0])

		var err error
		colNameAlpha, err = excelize.ColumnNumberToName(optColNum[0])
		if err == nil {
			colNameAlpha = "$" + colNameAlpha
		} else {
			colNameAlpha = ""
		}
	}

	sheet = strings.ToLower(sheet)
	name = strings.ToLower(name)

	// name:exact, sheet:exact
	for i, c := range cc {
		if strings.EqualFold(c.Name, name) && strings.EqualFold(c.Sheet, sheet) {
			return i
		}
	}
	for i, c := range cc {
		if strings.EqualFold(c.Name, colNameAlpha) && strings.EqualFold(c.Sheet, sheet) {
			return i
		}
	}
	for i, c := range cc {
		if strings.EqualFold(c.Name, colNameNum) && strings.EqualFold(c.Sheet, sheet) {
			return i
		}
	}

	// name:exact, sheet:wildcard
	for i, c := range cc {
		if strings.EqualFold(c.Name, name) && wildcardMatch(c.Sheet, sheet) {
			return i
		}
	}
	for i, c := range cc {
		if strings.EqualFold(c.Name, colNameAlpha) && wildcardMatch(c.Sheet, sheet) {
			return i
		}
	}
	for i, c := range cc {
		if strings.EqualFold(c.Name, colNameNum) && wildcardMatch(c.Sheet, sheet) {
			return i
		}
	}

	// name:exact, sheet:empty
	for i, c := range cc {
		if strings.EqualFold(c.Name, name) && strings.EqualFold(c.Sheet, "") {
			return i
		}
	}
	for i, c := range cc {
		if strings.EqualFold(c.Name, colNameAlpha) && strings.EqualFold(c.Sheet, "") {
			return i
		}
	}
	for i, c := range cc {
		if strings.EqualFold(c.Name, colNameNum) && strings.EqualFold(c.Sheet, "") {
			return i
		}
	}

	// name:wildcard, sheet:exact
	for i, c := range cc {
		if wildcardMatch(c.Name, name) && strings.EqualFold(c.Sheet, sheet) {
			return i
		}
	}

	// name:wildcard, sheet:wildcard
	for i, c := range cc {
		if wildcardMatch(c.Name, name) && wildcardMatch(c.Sheet, sheet) {
			return i
		}
	}

	// name:wildcard, sheet:empty
	for i, c := range cc {
		if wildcardMatch(c.Name, name) && strings.EqualFold(c.Sheet, "") {
			return i
		}
	}

	return -1
}

func wildcardMatch(pattern, name string) bool {
	if pattern == "*" {
		return true
	}
	if matched, _ := filepath.Match(pattern, name); matched {
		return true
	}
	return false
}

type colType string

const (
	typeUnknown  colType = ""
	typeText     colType = "text"
	typeNumber   colType = "number"
	typeDate     colType = "date"
	typeTime     colType = "time"
	typeDatetime colType = "datetime"
	typeBool     colType = "bool"
)

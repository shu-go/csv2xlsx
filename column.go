package main

import "strings"

type columns []column

type column struct {
	Name        string  `json:"name,omitempty"`
	Type        colType `json:"type,omitempty"`
	InputFormat string  `json:"format,omitempty"`
}

func (cc columns) findByName(name string) int {
	if cc == nil {
		return -1
	}

	for i := range cc {
		if strings.EqualFold(cc[i].Name, name) {
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

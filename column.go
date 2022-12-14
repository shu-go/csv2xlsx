package main

import "strings"

type Columns []Column

type Column struct {
	Name        string  `json:"name,omitempty"`
	Type        ColType `json:"type,omitempty"`
	InputFormat string  `json:"format,omitempty"`
}

func (cc Columns) FindByName(name string) int {
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

type ColType string

const (
	TypeUnknown  ColType = ""
	TypeText             = "text"
	TypeNumber           = "number"
	TypeDate             = "date"
	TypeTime             = "time"
	TypeDatetime         = "datetime"
)

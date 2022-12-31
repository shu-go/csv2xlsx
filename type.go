package main

import (
	"fmt"
	"regexp"
	"strings"
)

type baseType string

const (
	typeUnknown  baseType = ""
	typeText     baseType = "text"
	typeNumber   baseType = "number"
	typeDate     baseType = "date"
	typeTime     baseType = "time"
	typeDatetime baseType = "datetime"
	typeBool     baseType = "bool"
)

func (t baseType) derive(explicitInputFormat string) derivedType {
	derived := derivedType{
		baseType: t,
	}

	implicit, found := implicitInputFormats[t]
	if found {
		derived.implicitInputFormat = implicit
	}

	derived.explicitInputFormat = explicitInputFormat

	return derived
}

type derivedType struct {
	baseType baseType

	explicitInputFormat string
	implicitInputFormat string
}

var implicitInputFormats map[baseType]string

func parseType(s string) (derivedType, error) {
	declRE := regexp.MustCompile(`(text|number|datetime|date|time|bool)(?:\((.+)\))?`)
	subs := declRE.FindStringSubmatch(s)
	if subs == nil {
		return derivedType{}, fmt.Errorf("invalid type declaration %q", s)
	}

	base := baseType(subs[1])
	derived := base.derive(strings.TrimSpace(subs[2]))

	return derived, nil
}

func initImplicitDecls(dil, til, dtil string) {
	implicitInputFormats = make(map[baseType]string)

	implicitInputFormats[typeDate] = dil
	implicitInputFormats[typeTime] = til
	implicitInputFormats[typeDatetime] = dtil
}

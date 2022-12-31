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

func (t baseType) derive(explicitInputFormat, explicitOutputFormat string) derivedType {
	derived := derivedType{
		baseType: t,
	}

	implicit, found := implicitInputFormats[t]
	if found {
		derived.implicitInputFormat = implicit
	}

	implicit, found = implicitOutputFormats[t]
	if found {
		derived.implicitOutputFormat = implicit
	}

	derived.explicitInputFormat = explicitInputFormat
	derived.explicitOutputFormat = explicitOutputFormat

	return derived
}

type derivedType struct {
	baseType baseType

	explicitInputFormat string
	implicitInputFormat string

	explicitOutputFormat string
	implicitOutputFormat string
}

func (t derivedType) String() string {
	s := string(t.baseType)
	if t.explicitInputFormat != "" || t.implicitInputFormat != "" || t.explicitOutputFormat != "" || t.implicitOutputFormat != "" {
		s += "("

		if t.explicitInputFormat != "" && t.implicitInputFormat != "" {
			s += t.explicitInputFormat + " or " + t.implicitInputFormat
		} else if t.explicitInputFormat != "" {
			s += t.explicitInputFormat
		} else if t.implicitInputFormat != "" {
			s += t.implicitInputFormat
		}

		if t.explicitOutputFormat != "" || t.implicitOutputFormat != "" {
			s += "->"
		}
		if t.explicitOutputFormat != "" {
			s += t.explicitOutputFormat
		} else if t.implicitOutputFormat != "" {
			s += t.implicitOutputFormat
		}

		s += ")"
	}

	return s
}

var implicitInputFormats map[baseType]string
var implicitOutputFormats map[baseType]string

func parseType(s string) (derivedType, error) {
	declRE := regexp.MustCompile(`(text|number|datetime|date|time|bool)(?:\((.+?)(?:->(.+))?\))?`)
	subs := declRE.FindStringSubmatch(s)
	if subs == nil {
		return derivedType{}, fmt.Errorf("invalid type declaration %q", s)
	}

	base := baseType(subs[1])
	derived := base.derive(strings.TrimSpace(subs[2]), strings.TrimSpace(subs[3]))

	return derived, nil
}

func initImplicitDecls(dil, dol, til, tol, dtil, dtol string) {
	implicitInputFormats = make(map[baseType]string)
	implicitInputFormats[typeDate] = dil
	implicitInputFormats[typeTime] = til
	implicitInputFormats[typeDatetime] = dtil

	implicitOutputFormats = make(map[baseType]string)
	implicitOutputFormats[typeDate] = dol
	implicitOutputFormats[typeTime] = tol
	implicitOutputFormats[typeDatetime] = dtol
}

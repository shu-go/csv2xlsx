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

func (t baseType) derive(explicitInputLayout string) derivedType {
	derived := derivedType{
		baseType: t,
	}

	implicit, found := implicitInputLayouts[t]
	if found {
		derived.implicitInputLayout = implicit
	}

	derived.explicitInputLayout = explicitInputLayout

	return derived
}

type derivedType struct {
	baseType baseType

	explicitInputLayout string
	implicitInputLayout string
}

var implicitInputLayouts map[baseType]string

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
	implicitInputLayouts = make(map[baseType]string)

	implicitInputLayouts[typeDate] = dil
	implicitInputLayouts[typeTime] = til
	implicitInputLayouts[typeDatetime] = dtil
}

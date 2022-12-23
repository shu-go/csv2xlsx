package main

import (
	"bytes"
	"testing"

	"github.com/shu-go/gli"
	"github.com/shu-go/gotwant"
	"github.com/xuri/excelize/v2"
)

func dummyCmd() *globalCmd {
	app := gli.NewWith(&globalCmd{})
	icmd, _, _ := app.Parse([]string{"-o", "dummy", "dummyarg"})
	return icmd.(*globalCmd)
}

func newInput(name, content string) input {
	return input{Name: name, Reader: bytes.NewBufferString(content)}
}

func testValue(t *testing.T, oc outputContext, sheet, axis, want string, wantType ...excelize.CellType) {
	t.Helper()

	got, err := oc.output.GetCellValue(sheet, axis)
	gotwant.TestError(t, err, nil)
	gotwant.Test(t, got, want)

	if len(wantType) > 0 {
		gotType, err := oc.output.GetCellType(sheet, axis)
		gotwant.TestError(t, err, nil)
		gotwant.Test(t, gotType, wantType[0])
	}
}

func TestAis1(t *testing.T) {
	cmd := dummyCmd()

	oc, err := cmd.makeOutputContext(excelize.NewFile(), false)
	gotwant.TestError(t, err, nil)
	oc.inputs = []input{
		newInput("test.csv", "a\n1"),
	}

	err = cmd.convert(oc)
	gotwant.TestError(t, err, nil)

	testValue(t, oc, "test.csv", "A1", "a")
	testValue(t, oc, "test.csv", "A2", "1")
}

func TestGuess(t *testing.T) {
	cmd := dummyCmd()

	oc, err := cmd.makeOutputContext(excelize.NewFile(), false)
	gotwant.TestError(t, err, nil)
	oc.inputs = []input{
		newInput("test.csv", "a,b,c\n01,11,20220101"),
	}

	err = cmd.convert(oc)
	gotwant.TestError(t, err, nil)

	testValue(t, oc, "test.csv", "A2", "01")
	testValue(t, oc, "test.csv", "B2", "11")
	testValue(t, oc, "test.csv", "C2", "2022/01/01")

}

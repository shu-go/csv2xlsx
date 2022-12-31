package main

import (
	"bytes"
	"log"
	"math/rand"
	"strings"
	"testing"

	"github.com/shu-go/gli"
	"github.com/shu-go/gotwant"
	"github.com/xuri/excelize/v2"
)

func dummyCmd(args ...string) *globalCmd {
	appargs := append([]string{"-o", "dummy"}, args...)
	appargs = append(appargs, "dummycsvfile")

	app := gli.NewWith(&globalCmd{})
	icmd, _, err := app.Parse(appargs)
	if err != nil {
		log.Println(err)
	}

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
	tst := func(content, value string, args ...string) {
		t.Helper()

		cmd := dummyCmd(append([]string{"--header=0"}, args...)...)
		oc, err := cmd.makeOutputContext(excelize.NewFile(), false)
		gotwant.TestError(t, err, nil)
		oc.inputs = []input{
			newInput("test.csv", content),
		}

		err = cmd.convert(oc)
		gotwant.TestError(t, err, nil)
		testValue(t, oc, "test.csv", "A1", value)
	}

	t.Run("nohints", func(t *testing.T) {
		tst("01", "01")
		tst("11", "11")
		tst("20220101", "2022/01/01")
		tst("123456", "12:34:56")
		tst("true", "TRUE")
		tst("false", "FALSE")
	})

	t.Run("GuessNumber", func(t *testing.T) {
		tst("01", "01")
		tst("01", "1", "--columns", "#1:number")
	})

	t.Run("GuessBool", func(t *testing.T) {
		tst("1", "1")
		tst("1", "TRUE", "--columns", "#1:bool")
		tst("true", "TRUE")
		tst("false", "FALSE")
		tst("01", "01", "--columns", "#1:bool")
	})

	t.Run("GuessDate", func(t *testing.T) {
		tst("20220304", "2022/03/04")
		tst("4-3-22", "2022/03/04", "--columns", "#1:date(d-m-y)")
		tst("4-3-22", "2022-03-04", "--columns", "#1:date(d-m-y -> yyyy-mm-dd)")
		tst("4-3-22", "2022/03/04", "--columns", "#1:date(2-1-06)")
		tst("4-3-22", "4-3-22", "--columns", "#1:date(dd-mm-yyyy)") // failure
		tst("Feb 4 2008", "2008/02/04", "--columns", "#1:date(Jan 2 2006)")
	})

	t.Run("GuessTime", func(t *testing.T) {
		tst("123456", "12:34:56")
		tst("012345", "012345") // failure
		tst("012345", "01:23:45", "--columns", "#1:time")
		tst("1 2 3", "01:02:03", "--columns", "#1:time(h m s)")
		tst("1 2 3", "1 2 3", "--columns", "#1:time(hh mm ss)") // failure
	})

	t.Run("GuessDatetime", func(t *testing.T) {
		tst("20220304", "20220304", "--columns", "#1:datetime")
		tst("20220304 123456", "2022/03/04 12:34:56", "--columns", "#1:datetime")
		tst("Feb 4, 2008 4:45pm", "2008/02/04 16:45:00", "--columns", `#1:datetime(Jan 2\, 2006 3:04pm)`, "-d", ";")
		tst("Feb 4, 2008 4:45pm", "2008 02 04 16 45 00", "--columns", `#1:datetime(Jan 2\, 2006 3:04pm -> yyyy mm dd hh mm ss)`, "-d", ";")
		tst("true", "true", "--columns", "#1:datetime")
	})
}

func TestMultiple(t *testing.T) {
	cmd := dummyCmd([]string{"--header=0"}...)
	oc, err := cmd.makeOutputContext(excelize.NewFile(), false)
	gotwant.TestError(t, err, nil)
	oc.inputs = []input{
		newInput("test.csv", "ichi"),
		newInput("test2.csv", "ni"),
	}
	err = cmd.convert(oc)
	gotwant.TestError(t, err, nil)
	testValue(t, oc, "test.csv", "A1", "ichi")
	testValue(t, oc, "test2.csv", "A1", "ni")

	//

	oc.inputs = []input{
		newInput("test.csv", "ichi"),
		newInput("test3.csv", "san\nthree"),
	}
	err = cmd.convert(oc)
	gotwant.TestError(t, err, nil)
	testValue(t, oc, "test.csv", "A1", "ichi")
	testValue(t, oc, "test2.csv", "A1", "ni")
	testValue(t, oc, "test3.csv", "A1", "san")
	testValue(t, oc, "test3.csv", "A2", "three")

	//

	oc.inputs = []input{
		newInput("test3.csv", "san"),
	}
	err = cmd.convert(oc)
	gotwant.TestError(t, err, nil)
	testValue(t, oc, "test.csv", "A1", "ichi")
	testValue(t, oc, "test2.csv", "A1", "ni")
	testValue(t, oc, "test3.csv", "A1", "san")
	testValue(t, oc, "test3.csv", "A2", "")
}

func BenchmarkGuess(b *testing.B) {
	tst := func(content, value string, args ...string) {
		b.Helper()

		cmd := dummyCmd(append([]string{"--header=0"}, args...)...)
		oc, _ := cmd.makeOutputContext(excelize.NewFile(), false)
		oc.inputs = []input{
			newInput("test.csv", content),
		}

		_ = cmd.convert(oc)
	}

	gen := func(n int) []string {
		templ := []string{
			"abcdefgh",
			"01234567",
			"12345678",
			"20221231",
			"012345",
			"123456",
			"20221231 012345",
			"20221231 123456",
			"true",
			"false",
		}
		l := len(templ)

		s := make([]string, n)
		for i := range s {
			s[i] = templ[rand.Intn(l)]
		}

		return s
	}

	b.Run("10", func(b *testing.B) {
		s := strings.Join(gen(00), "\n")
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			tst(s, "")
		}
	})

	b.Run("100", func(b *testing.B) {
		s := strings.Join(gen(100), "\n")
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			tst(s, "")
		}
	})

	b.Run("1000", func(b *testing.B) {
		s := strings.Join(gen(1000), "\n")
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			tst(s, "")
		}
	})

	b.Run("10000", func(b *testing.B) {
		s := strings.Join(gen(10000), "\n")
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			tst(s, "")
		}
	})
}

package main

import (
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"encoding/csv"

	"github.com/andrew-d/go-termutil"
	"github.com/shu-go/gli/v2"
	"github.com/xuri/excelize/v2"
)

// Version is app version
var Version string

type globalCmd struct {
	Output    string `cli:"output,o=FILENAME" required:"true"`
	Delimiter string `cli:"d" default:"," help:"a value delimiter"`

	Header int `cli:"header" default:"1" help:"-1 when no header"`

	GuessType bool `cli:"guess,g" default:"true" help:"guess cell type by --columns or CSV values"`

	DateFmt     string `cli:"date,df" default:"ymd" help:"global input format of date over columns"`
	TimeFmt     string `cli:"time,tf" default:"hms" help:"global input format of time over columns"`
	DatetimeFmt string `cli:"datetime,dtf" default:"20060102 150405" help:"global input format of datetime over columns"`

	DateXlsxFmt     string `cli:"date-xlsx,dxf" default:"yyyy/mm/dd" help:"global output format of date over columns"`
	TimeXlsxFmt     string `cli:"time-xlsx,txf" default:"hh:mm:ss" help:"global output format of time over columns"`
	DatetimeXlsxFmt string `cli:"datetime-xlsx,dtxf" default:"yyyy/mm/dd hh:mm:ss" help:"global output format of datetime over columns"`

	NumberXlsxFmt string `cli:"number-xlsx,nxf" default:""`

	Columns map[string]string `cli:"columns,cols" help:"[SHEET!]COLUMN_NAME:TYPE[(INPUT_FORMAT[->OUTPUT_FORMAT])],..."`

	PipelinedName string `cli:"pipelined-name,name=SHEET_NAME" help:"the name of a pipelined CSV" default:"Sheet1"`
}

func (c globalCmd) Before(args []string) error {
	if c.Output == "" {
		return errors.New("--output is required")
	}

	if termutil.Isatty(os.Stdin.Fd()) && len(args) == 0 {
		return errors.New("at least one csv file is required")
	}

	return nil
}

func (c globalCmd) Run(args []string) error {
	var xlsxfile *excelize.File
	exists := false
	if fileExists(c.Output) {
		exists = true
		f, err := excelize.OpenFile(filepath.Clean(c.Output))
		if err != nil {
			return err
		}
		xlsxfile = f
	} else {
		xlsxfile = excelize.NewFile()
	}

	oc, err := c.makeOutputContext(xlsxfile, exists)
	if err != nil {
		return err
	}

	if !termutil.Isatty(os.Stdin.Fd()) {
		oc.inputs = append(oc.inputs, input{
			Name:   c.PipelinedName,
			Reader: os.Stdin,
		})
	}

	for _, csvfilename := range args {
		csvfilename = filepath.Clean(csvfilename)

		f, err := os.Open(csvfilename)
		if err != nil {
			return err
		}
		defer func(f *os.File) { f.Close() }(f)

		oc.inputs = append(oc.inputs, input{
			Name:   csvfilename,
			Reader: f,
		})
	}

	err = c.convert(oc)
	if err != nil {
		return err
	}

	err = xlsxfile.SaveAs(c.Output)
	if err != nil {
		return err
	}

	return nil
}

type input struct {
	Name   string
	Reader io.Reader
}

type outputContext struct {
	output      *excelize.File
	overwriting bool

	inputs []input

	hints columns

	styles map[string]int
}

func (c globalCmd) makeOutputContext(xlsxfile *excelize.File, overwriting bool) (outputContext, error) {
	oc := outputContext{
		output:      xlsxfile,
		overwriting: overwriting,
		styles:      make(map[string]int),
	}

	for k, v := range c.Columns {
		typ, err := parseType(v)
		if err != nil {
			return outputContext{}, err
		}

		oc.hints = append(oc.hints, newColumn(k, typ))
	}
	/*
		for _, h := range oc.hints {
			log.Println(h)
		}
	*/

	style, err := defineStyle(xlsxfile, c.DateXlsxFmt)
	if err != nil {
		return outputContext{}, err
	}
	oc.styles[c.DateXlsxFmt] = style

	style, err = defineStyle(xlsxfile, c.TimeXlsxFmt)
	if err != nil {
		return outputContext{}, err
	}
	oc.styles[c.TimeXlsxFmt] = style

	style, err = defineStyle(xlsxfile, c.DatetimeXlsxFmt)
	if err != nil {
		return outputContext{}, err
	}
	oc.styles[c.DatetimeXlsxFmt] = style

	style, err = defineStyle(xlsxfile, c.NumberXlsxFmt)
	if err != nil {
		return outputContext{}, err
	}
	oc.styles[c.NumberXlsxFmt] = style

	return oc, nil
}

func (c globalCmd) convert(oc outputContext) error {
	initImplicitDecls(c.DateFmt, c.DateXlsxFmt, c.TimeFmt, c.TimeXlsxFmt, c.DatetimeFmt, c.DatetimeXlsxFmt, c.NumberXlsxFmt)

	for _, in := range oc.inputs {
		err := c.convertOne(oc, in.Name, in.Reader)
		if err != nil {
			return err
		}
	}

	sheet1 := false
	for _, in := range oc.inputs {
		sheet1 = sheet1 || strings.EqualFold(in.Name, "Sheet1")
	}

	if !sheet1 && !oc.overwriting {
		oc.output.DeleteSheet("Sheet1")
	}

	oc.output.SetActiveSheet(0)

	return nil
}

func (c globalCmd) tempSheetName(oc outputContext) string {
	name := "_CSV2XLSX_TEMP_"
	for {
		if idx, _ := oc.output.GetSheetIndex(name); idx == -1 {
			return name
		}

		name += "a"
	}
}

func (c globalCmd) convertOne(oc outputContext, sheet string, input io.Reader) error {
	tempname := c.tempSheetName(oc)
	oc.output.NewSheet(tempname)
	oc.output.DeleteSheet(sheet)
	oc.output.NewSheet(sheet)
	oc.output.DeleteSheet(tempname)

	r := csv.NewReader(input)
	if len(c.Delimiter) > 0 {
		r.Comma = []rune(c.Delimiter)[0]
	}

	csvrindex := 0
	xlsxrindex := 0
	columns := []string{}

	for {
		fields, err := r.Read()
		if err == io.EOF {
			break
		} else if err != nil {
			return err
		}

		if csvrindex == c.Header-1 {
			for cindex := range fields {
				columns = append(columns, strings.TrimSpace(fields[cindex]))
			}

			err := writeXlsxHeader(oc.output, sheet, xlsxrindex, fields)
			if err != nil {
				return err
			}

			xlsxrindex++
		}
		if csvrindex <= c.Header-1 {
			csvrindex++
			continue
		}

		for i := len(columns); i < len(fields); i++ {
			colName, err := excelize.ColumnNumberToName(i + 1)
			if err != nil {
				return err
			}

			columns = append(columns, "$"+colName)
		}

		for cindex, value := range fields {
			colName := columns[cindex]

			addr, err := excelize.CoordinatesToCellName(cindex+1, xlsxrindex+1)
			if err != nil {
				return fmt.Errorf("%v: %v\n", colName, err)
			}

			if len(value) == 0 {
				continue
			}

			if !c.GuessType {
				err = oc.output.SetCellValue(sheet, addr, value)
				if err != nil {
					return err
				}

				continue
			}

			hindex := oc.hints.findByName(sheet, colName, cindex+1)
			col := column{}
			if hindex != -1 {
				col = oc.hints[hindex]
			}

			typ, ival := c.guess(value, col)

			err = writeXlsx(oc.output, sheet, addr, typ, ival, oc.styles)
			if err != nil {
				return err
			}
		}

		xlsxrindex++
		csvrindex++
	}

	return nil
}

func writeXlsxHeader(f *excelize.File, sheet string, rindex int, fields []string) error {
	for cindex, value := range fields {
		addr, err := excelize.CoordinatesToCellName(cindex+1, rindex+1)
		if err != nil {
			return err
		}
		err = f.SetCellValue(sheet, addr, value)
		if err != nil {
			return err
		}
	}
	return nil
}

func writeXlsx(f *excelize.File, sheet string, axis string, typ derivedType, value interface{}, styles map[string]int) error {
	outputfmt := typ.explicitOutputFormat
	if outputfmt == "" {
		outputfmt = typ.implicitOutputFormat
	}
	style, found := styles[outputfmt]
	if !found {
		var err error
		style, err = defineStyle(f, outputfmt)
		if err != nil {
			return err
		}

	}

	switch typ.baseType {
	case typeText:
		err := f.SetCellValue(sheet, axis, value)
		if err != nil {
			return err
		}

	case typeDatetime:
		err := setCellValueAndStyle(f, sheet, axis, value, style)
		if err != nil {
			return err
		}

	case typeDate:
		err := setCellValueAndStyle(f, sheet, axis, value, style)
		if err != nil {
			return err
		}

	case typeTime:
		tval := value.(time.Time)
		if y, m, d := tval.Date(); y == 0 && m == 1 && d == 1 {
			tval = time.Date(1900, 1, 1, tval.Hour(), tval.Minute(), tval.Second(), tval.Nanosecond(), tval.Location())
		}
		err := setCellValueAndStyle(f, sheet, axis, tval, style)
		if err != nil {
			return err
		}

	case typeNumber:
		err := setCellValueAndStyle(f, sheet, axis, value, style)
		if err != nil {
			return err
		}

	case typeFormula:
		err := setCellFormulaAndStyle(f, sheet, axis, value, style)
		if err != nil {
			return err
		}

	default:
		err := f.SetCellValue(sheet, axis, value)
		if err != nil {
			return err
		}
	}

	return nil
}

func (c globalCmd) guess(value string, col column) (derivedType, interface{}) {
	if typ, ival := c.guessByColType(value, col); typ.baseType != typeUnknown {
		return typ, ival
	}

	if value[0] == '\'' || value[0] == '0' {
		return typeText.derive("", ""), value
	}
	if value[0] == '=' {
		return typeFormula.derive("", ""), value
	}
	if strings.ToLower(value) == "true" {
		return typeBool.derive("", ""), true
	}
	if strings.ToLower(value) == "false" {
		return typeBool.derive("", ""), false
	}

	typetest := typeDatetime.derive("", "")
	if t, ok := parseTime(value, typetest.implicitInputFormat); ok {
		return typetest, t
	}

	typetest = typeDate.derive("", "")
	ptns := translateDatePatterns(typetest.implicitInputFormat)
	if t, ok := parseTime(value, ptns...); ok {
		return typetest, t
	}

	typetest = typeTime.derive("", "")
	ptns = translateTimePatterns(typetest.implicitInputFormat)
	if t, ok := parseTime(value, ptns...); ok {
		return typetest, t
	}

	if f, err := strconv.ParseFloat(value, 64); err == nil {
		return typeNumber.derive("", ""), f
	}

	return typeUnknown.derive("", ""), value
}

func (c globalCmd) guessByColType(value string, col column) (derivedType, interface{}) {
	switch col.Type.baseType {
	case typeText:
		return col.Type, value

	case typeNumber:
		if f, err := strconv.ParseFloat(value, 64); err == nil {
			return col.Type, f
		}

	case typeDate:
		ptns := translateDatePatterns(col.Type.explicitInputFormat)
		ptns = append(ptns, translateDatePatterns(col.Type.implicitInputFormat)...)
		if t, ok := parseTime(value, ptns...); ok {
			return col.Type, t
		}

	case typeTime:
		ptns := translateTimePatterns(col.Type.explicitInputFormat)
		ptns = append(ptns, translateTimePatterns(col.Type.implicitInputFormat)...)
		if t, ok := parseTime(value, ptns...); ok {
			return col.Type, t
		}

	case typeDatetime:
		ptns := append([]string{}, col.Type.explicitInputFormat)
		ptns = append(ptns, col.Type.implicitInputFormat)
		if t, ok := parseTime(value, ptns...); ok {
			return col.Type, t
		}

	case typeBool:
		if b, err := strconv.ParseBool(value); err == nil {
			return col.Type, b
		}

	case typeFormula:
		return col.Type, value

	default: // nop
	}

	if col.Type.baseType != typeUnknown {
		return typeText.derive("", ""), value
	}

	return typeUnknown.derive("", ""), value
}

func parseTime(value string, layouts ...string) (time.Time, bool) {
	for i := range layouts {
		if len(layouts[i]) != len(value) {
			continue
		}

		if t, err := time.Parse(layouts[i], value); err == nil {
			return t, true
		}
	}
	return time.Time{}, false
}

func translateDatePatterns(ptn string) []string {
	if ptn == "" {
		return nil
	}

	var ptns []string
	if strings.ContainsAny(ptn, "ymd") {
		if !strings.Contains(ptn, "yyyy") &&
			!strings.Contains(ptn, "mm") &&
			!strings.Contains(ptn, "dd") {
			//
			p := ptn
			p = strings.ReplaceAll(p, "yy", "2006")
			p = strings.ReplaceAll(p, "y", "2006")
			p = strings.ReplaceAll(p, "m", "01")
			p = strings.ReplaceAll(p, "d", "02")
			ptns = append(ptns, p)

			p = ptn
			p = strings.ReplaceAll(p, "yy", "06")
			p = strings.ReplaceAll(p, "y", "06")
			p = strings.ReplaceAll(p, "m", "01")
			p = strings.ReplaceAll(p, "d", "02")
			ptns = append(ptns, p)

			p = ptn
			p = strings.ReplaceAll(p, "yy", "06")
			p = strings.ReplaceAll(p, "y", "06")
			p = strings.ReplaceAll(p, "m", "1")
			p = strings.ReplaceAll(p, "d", "2")
			ptns = append(ptns, p)
		} else {
			p := ptn
			p = strings.ReplaceAll(p, "yyyy", "2006")
			p = strings.ReplaceAll(p, "yy", "06")
			p = strings.ReplaceAll(p, "y", "06")
			p = strings.ReplaceAll(p, "mm", "01")
			p = strings.ReplaceAll(p, "m", "1")
			p = strings.ReplaceAll(p, "dd", "02")
			p = strings.ReplaceAll(p, "d", "2")
			ptns = append(ptns, p)
		}
	} else {
		ptns = append(ptns, ptn)
	}
	return ptns
}

func translateTimePatterns(ptn string) []string {
	var ptns []string
	if strings.ContainsAny(ptn, "hms") {
		if !strings.Contains(ptn, "hh") &&
			!strings.Contains(ptn, "mm") &&
			!strings.Contains(ptn, "ss") {
			//
			p := ptn
			p = strings.ReplaceAll(p, "hh", "15")
			p = strings.ReplaceAll(p, "h", "15")
			p = strings.ReplaceAll(p, "m", "04")
			p = strings.ReplaceAll(p, "s", "05")
			ptns = append(ptns, p)

			p = ptn
			p = strings.ReplaceAll(p, "h", "3")
			p = strings.ReplaceAll(p, "m", "4")
			p = strings.ReplaceAll(p, "s", "5")
			ptns = append(ptns, p)
		} else {
			p := ptn
			p = strings.ReplaceAll(p, "hh", "15")
			p = strings.ReplaceAll(p, "h", "15")
			p = strings.ReplaceAll(p, "mm", "04")
			p = strings.ReplaceAll(p, "m", "4")
			p = strings.ReplaceAll(p, "ss", "05")
			p = strings.ReplaceAll(p, "s", "5")
			ptns = append(ptns, p)
		}
	} else {
		ptns = append(ptns, ptn)
	}
	return ptns
}

func defineStyle(f *excelize.File, s string) (int, error) {
	if s == "" {
		return f.NewStyle(&excelize.Style{NumFmt: 0})
	}

	return f.NewStyle(&excelize.Style{CustomNumFmt: &s})
}

func setCellValueAndStyle(f *excelize.File, sheet, axis string, value interface{}, styleID int) error {
	err := f.SetCellValue(sheet, axis, value)
	if err != nil {
		return err
	}

	err = f.SetCellStyle(sheet, axis, axis, styleID)
	if err != nil {
		return err
	}

	return nil
}

func setCellFormulaAndStyle(f *excelize.File, sheet, axis string, formula interface{}, styleID int) error {
	var fstr string
	if s, ok := formula.(string); ok {
		fstr = s
	} else if s, ok := formula.(fmt.Stringer); ok {
		fstr = s.String()
	}

	err := f.SetCellFormula(sheet, axis, fstr)
	if err != nil {
		return err
	}

	err = f.SetCellStyle(sheet, axis, axis, styleID)
	if err != nil {
		return err
	}

	return nil
}

func fileExists(name string) bool {
	info, err := os.Stat(filepath.Clean(name))
	return err == nil && !info.IsDir()
}

func main() {
	app := gli.NewWith(&globalCmd{})
	app.Name = "csv2xlsx"
	app.Desc = "CSV to XLSX file converter"
	app.Version = Version
	app.Usage = `csv2xlsx [options] -o FILENAME CSV_FILENAME [CSV_FILENAME...]

--columns [SHEET!]COLUMN_NAME:TYPE[(INPUT_FORMAT[->OUTPUT_FORMAT])]
  SHEET = CSV_FILENAME
  TYPE = text | number | date | time | datetime | bool | formula
  INPUT_FORMAT
    date: yyyy, yy, y, 2006, 06, mm, m, 01, 1, dd, d, 02, 2
    time: hh, h, 15, 3, mm, m, 04, 4, ss, s, 05, 5
    datetime: 2006, 06, 01, 1, 02, 2, 15, 3, 04, 4, 05, 5
  Examples:
    csv2xlsx -o dest.xlsx src.csv
    csv2xlsx -o dest.xlsx --columns num_*:number src.csv
    csv2xlsx -o dest.xlsx --columns num_*:"number(->#\,##0.00)" src.csv
`
	app.Copyright = "(C) 2022 Shuhei Kubota"
	err := app.Run(os.Args)
	if err != nil {
		os.Exit(1)
	}

}

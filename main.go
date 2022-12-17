package main

import (
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	"encoding/csv"

	"github.com/shu-go/gli"
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

	Columns gli.Map `cli:"columns,cols" help:""`
}

func (c globalCmd) Run(args []string) error {
	if c.Output == "" {
		return errors.New("--output is required")
	}

	if len(args) == 0 {
		return errors.New("at least one csv file is required")
	}

	var columnHints columns
	hintRE := regexp.MustCompile(`(text|number|date|time|datetime)(:?\((.+)\))?`)
	for k, v := range c.Columns {
		subs := hintRE.FindStringSubmatch(v)
		if subs == nil {
			return fmt.Errorf("a value of --columns is invalid: %q:%q", k, v)
		}

		col := newColumn(k, colType(subs[1]), strings.TrimSpace(subs[2]))
		if col.Type == typeDate && col.InputFormat == "" {
			col.InputFormat = c.DateFmt
		} else if col.Type == typeTime && col.InputFormat == "" {
			col.InputFormat = c.TimeFmt
		} else if col.Type == typeDatetime && col.InputFormat == "" {
			col.InputFormat = c.DatetimeFmt
		}

		columnHints = append(columnHints, col)
	}

	var datePtns []string = translateDatePatterns(c.DateFmt)
	var timePtns []string = translateTimePatterns(c.TimeFmt)

	xlsxfile := excelize.NewFile()

	dateStyle, err := defineStyle(xlsxfile, c.DateXlsxFmt)
	if err != nil {
		return err
	}
	timeStyle, err := defineStyle(xlsxfile, c.TimeXlsxFmt)
	if err != nil {
		return err
	}
	datetimeStyle, err := defineStyle(xlsxfile, c.DatetimeXlsxFmt)
	if err != nil {
		return err
	}
	numberStyle, err := defineStyle(xlsxfile, c.NumberXlsxFmt)
	if err != nil {
		return err
	}

	for _, csvfilename := range args {
		ctx := outputContext{
			xlsxfile:    xlsxfile,
			csvfilename: csvfilename,

			hints: columnHints,

			dateStyle:     dateStyle,
			timeStyle:     timeStyle,
			datetimeStyle: datetimeStyle,
			numberStyle:   numberStyle,

			datePtns: datePtns,
			timePtns: timePtns,
		}
		err = c.runOneCSV(ctx)
		if err != nil {
			return err
		}
	}

	xlsxfile.DeleteSheet("Sheet1")
	xlsxfile.SetActiveSheet(0)
	err = xlsxfile.SaveAs(c.Output)
	if err != nil {
		return err
	}

	return nil
}

type outputContext struct {
	xlsxfile    *excelize.File
	csvfilename string

	hints columns

	dateStyle, timeStyle, datetimeStyle, numberStyle int

	datePtns, timePtns []string
}

func (c globalCmd) runOneCSV(oc outputContext) error {
	sheet := filepath.Base(oc.csvfilename)

	csvfile, err := os.Open(oc.csvfilename)
	if err != nil {
		return err
	}
	defer csvfile.Close()

	oc.xlsxfile.NewSheet(sheet)

	r := csv.NewReader(csvfile)
	csvrindex := 0
	xlsxrindex := 0
	columns := columns{}

	for {
		fields, err := r.Read()
		if err == io.EOF {
			break
		} else if err != nil {
			return err
		}

		if csvrindex == c.Header-1 {
			for cindex := range fields {
				colname := strings.TrimSpace(fields[cindex])
				var col column
				if i := oc.hints.findByName(sheet, colname); i != -1 {
					col = oc.hints[i]
				} else {
					col = newColumn(sheet+"!"+colname, typeUnknown, "")
				}
				columns = append(columns, col)
			}

			err := writeXlsxHeader(oc.xlsxfile, sheet, xlsxrindex, fields)
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

			var col column
			if i := oc.hints.findByName(sheet, colName); i != -1 {
				col = oc.hints[i]
			} else {
				col = newColumn(sheet+"!"+colName, typeUnknown, "")
			}
			columns = append(columns, col)
		}

		for cindex, value := range fields {
			g := columns[cindex]

			addr, err := excelize.CoordinatesToCellName(cindex+1, xlsxrindex+1)
			if err != nil {
				return fmt.Errorf("%v: %v\n", g.Name, err)
			}

			if len(value) == 0 {
				continue
			}

			if !c.GuessType {
				err = oc.xlsxfile.SetCellValue(sheet, addr, value)
				if err != nil {
					return err
				}

				continue
			}

			typ, ival := c.guess(value, g, oc.datePtns, oc.timePtns)

			err = writeXlsx(oc.xlsxfile, sheet, addr, typ, ival, oc.numberStyle, oc.dateStyle, oc.timeStyle, oc.datetimeStyle)
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

func writeXlsx(f *excelize.File, sheet string, axis string, typ colType, value interface{}, numberStyle, dateStyle, timeStyle, datetimeStyle int) error {
	switch typ {
	case typeText:
		err := f.SetCellValue(sheet, axis, value)
		if err != nil {
			return err
		}

	case typeDatetime:
		err := setCellValueAndStyle(f, sheet, axis, value, datetimeStyle)
		if err != nil {
			return err
		}

	case typeDate:
		err := setCellValueAndStyle(f, sheet, axis, value, dateStyle)
		if err != nil {
			return err
		}

	case typeTime:
		tval := value.(time.Time)
		if y, m, d := tval.Date(); y == 0 && m == 1 && d == 1 {
			tval = time.Date(1900, 1, 1, tval.Hour(), tval.Minute(), tval.Second(), tval.Nanosecond(), tval.Location())
		}
		err := setCellValueAndStyle(f, sheet, axis, tval, timeStyle)
		if err != nil {
			return err
		}

	case typeNumber:
		err := setCellValueAndStyle(f, sheet, axis, value, numberStyle)
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

func (c globalCmd) guess(value string, col column, dateLayouts, timeLayouts []string) (colType, interface{}) {
	typ, ival := c.guessByColType(value, col)
	if typ != typeUnknown {
		return typ, ival
	}

	if value[0] == '\'' || value[0] == '0' {
		return typeText, value
	}
	if t, ok := parseTime(value, c.DatetimeFmt); ok {
		return typeDatetime, t
	}
	if t, ok := parseTime(value, dateLayouts...); ok {
		return typeDate, t
	}
	if t, ok := parseTime(value, timeLayouts...); ok {
		return typeTime, t
	}
	if f, err := strconv.ParseFloat(value, 64); err == nil {
		return typeNumber, f
	}

	return typeUnknown, value
}

func (c globalCmd) guessByColType(value string, col column) (colType, interface{}) {
	switch col.Type {
	case typeText:
		return typeText, value

	case typeNumber:
		if f, err := strconv.ParseFloat(value, 64); err == nil {
			return typeNumber, f
		}

	case typeDate:
		var ptns []string = translateDatePatterns(col.InputFormat)
		if t, ok := parseTime(value, ptns...); ok {
			return typeDate, t
		}

	case typeTime:
		var ptns []string = translateTimePatterns(col.InputFormat)
		if t, ok := parseTime(value, ptns...); ok {
			return typeTime, t
		}

	case typeDatetime:
		if t, ok := parseTime(value, col.InputFormat); ok {
			return typeDatetime, t
		}

	default: // nop
	}

	return typeUnknown, value
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

func main() {
	app := gli.NewWith(&globalCmd{})
	app.Name = "csv2xlsx"
	app.Desc = ""
	app.Version = Version
	app.Usage = `--columns COLUMN_NAME:TYPE(INPUT_FORMAT)
  TYPE = [text|number|date|time|datetime]
  INPUT_FORMAT
    date: yyyy, yy, y, 2006, 06, mm, m, 01, 1, dd, d, 02, 2
    time: hh, h, 15, 3, mm, m, 04, 4, ss, s, 05, 5
    datetime: 2006, 06, 01, 1, 02, 2, 15, 3, 04, 4, 05, 5
`
	app.Copyright = "(C) 2022 Shuhei Kubota"
	err := app.Run(os.Args)
	if err != nil {
		os.Exit(1)
	}

}

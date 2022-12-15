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

	//csv "github.com/JensRantil/go-csv"

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

	Columns gli.Map `cli:"columns,cols" help:""`
}

func (c globalCmd) Run(args []string) error {
	if c.Output == "" {
		return errors.New("--output is required")
	}

	var columnHints columns
	hintRE := regexp.MustCompile(`(text|number|date|time|datetime)(\((.+)\))?`)
	for k, v := range c.Columns {
		subs := hintRE.FindStringSubmatch(v)
		if subs == nil {
			return fmt.Errorf("a value of --columns is invalid: %q:%q", k, v)
		}

		col := column{Name: k, Type: colType(subs[1]), InputFormat: strings.TrimSpace(subs[3])}
		if col.Type == typeDate && col.InputFormat == "" {
			col.InputFormat = c.DateFmt
		} else if col.Type == typeTime && col.InputFormat == "" {
			col.InputFormat = c.TimeFmt
		} else if col.Type == typeDatetime && col.InputFormat == "" {
			col.InputFormat = c.DatetimeFmt
		}

		columnHints = append(columnHints, col)
	}

	x := excelize.NewFile()

	exp := c.DateXlsxFmt
	dateStyle, err := x.NewStyle(&excelize.Style{CustomNumFmt: &exp})
	if err != nil {
		return err
	}
	exp = c.TimeXlsxFmt
	timeStyle, err := x.NewStyle(&excelize.Style{CustomNumFmt: &exp})
	if err != nil {
		return err
	}
	exp = c.DatetimeXlsxFmt
	datetimeStyle, err := x.NewStyle(&excelize.Style{CustomNumFmt: &exp})
	if err != nil {
		return err
	}

	csvFileName := args[0]
	sheetName := filepath.Base(csvFileName)

	f, err := os.Open(csvFileName)
	if err != nil {
		return err
	}
	defer f.Close()

	x.NewSheet(sheetName)

	var datePtns []string = translateDatePatterns(c.DateFmt)
	var timePtns []string = translateTimePatterns(c.TimeFmt)

	r := csv.NewReader(f)
	csvrindex := 0
	xlsxrindex := 0
	columns := columns{}

	for {
		fields, err := r.Read()
		//log.Println(fields)
		if err == io.EOF {
			//log.Println("EOF")
			break
		} else if err != nil {
			return err
		}

		if csvrindex == c.Header-1 {
			for cindex, val := range fields {
				//log.Printf("%v:%v\n", cindex, val)
				col := column{Name: strings.TrimSpace(val)}
				if i := columnHints.findByName(col.Name); i != -1 {
					col = columnHints[i]
				}
				columns = append(columns, col)

				addr, err := excelize.CoordinatesToCellName(cindex+1, xlsxrindex+1)
				if err != nil {
					return err
				}
				err = x.SetCellValue(sheetName, addr, val)
				if err != nil {
					return err
				}
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

			col := column{Name: "#" + colName}
			if i := columnHints.findByName(col.Name); i != -1 {
				col = columnHints[i]
			}
			columns = append(columns, col)
		}

		for cindex, val := range fields {
			g := columns[cindex]
			//log.Print(g.Name, val, g.Type)

			addr, err := excelize.CoordinatesToCellName(cindex+1, xlsxrindex+1)
			if err != nil {
				return fmt.Errorf("%v: %v\n", g.Name, err)
			}
			//log.Println(addr, g.Name, val)

			if len(val) == 0 {
				continue
			}

			if !c.GuessType {
				err = x.SetCellValue(sheetName, addr, val)
				if err != nil {
					return err
				}

				continue
			}

			typ, ival := c.guess(val, g, datePtns, timePtns)

			//log.Println(addr, g.Name, val, typ)

			switch typ {
			case typeText:
				err = x.SetCellValue(sheetName, addr, val)
				if err != nil {
					return err
				}

			case typeDatetime:
				//log.Println(addr, tvalue)
				err = setCellValueAndStyle(x, sheetName, addr, ival, datetimeStyle)
				if err != nil {
					return err
				}

			case typeDate:
				//log.Println(addr, tvalue)
				err = setCellValueAndStyle(x, sheetName, addr, ival, dateStyle)
				if err != nil {
					return err
				}

			case typeTime:
				tval := ival.(time.Time)
				if y, m, d := tval.Date(); y == 0 && m == 1 && d == 1 {
					tval = time.Date(1900, 1, 1, tval.Hour(), tval.Minute(), tval.Second(), tval.Nanosecond(), tval.Location())
				}
				//log.Println(addr, tval)
				err = setCellValueAndStyle(x, sheetName, addr, tval, timeStyle)
				if err != nil {
					return err
				}

			case typeNumber:
				//log.Println(addr, fvalue)
				err = x.SetCellValue(sheetName, addr, ival)
				if err != nil {
					return err
				}

			default:
				err = x.SetCellValue(sheetName, addr, val)
				if err != nil {
					return err
				}
			}
		}

		xlsxrindex++
		csvrindex++
	}

	x.SetActiveSheet(1)
	err = x.SaveAs(c.Output)
	if err != nil {
		return err
	}

	return nil
}

func (c globalCmd) guess(value string, col column, dateLayouts, timeLayouts []string) (colType, interface{}) {
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

	if value[0] == '\'' {
		return typeText, value
	}
	if value[0] == '0' {
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

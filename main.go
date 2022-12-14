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

	var columnHints Columns
	hintRE := regexp.MustCompile(`(text|number|date|time|datetime)(\((.+)\))?`)
	if len(c.Columns) != 0 {
		for k, v := range c.Columns {
			subs := hintRE.FindStringSubmatch(v)
			if subs == nil {
				return fmt.Errorf("a value of --columns is invalid: %q:%q", k, v)
			}

			col := Column{Name: k, Type: ColType(subs[1]), InputFormat: strings.TrimSpace(subs[3])}
			if col.Type == TypeDate && col.InputFormat == "" {
				col.InputFormat = c.DateFmt
			} else if col.Type == TypeTime && col.InputFormat == "" {
				col.InputFormat = c.TimeFmt
			} else if col.Type == TypeDatetime && col.InputFormat == "" {
				col.InputFormat = c.DatetimeFmt
			}

			columnHints = append(columnHints, col)
		}
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
	rindex := 0
	columns := Columns{}

	for {
		fields, err := r.Read()
		//log.Println(fields)
		if err == io.EOF {
			//log.Println("EOF")
			break
		} else if err != nil {
			return err
		}

		if rindex == c.Header-1 {
			for cindex, val := range fields {
				//log.Printf("%v:%v\n", cindex, val)
				col := Column{Name: strings.TrimSpace(val)}
				if i := columnHints.FindByName(col.Name); i != -1 {
					col = columnHints[i]
				}
				columns = append(columns, col)

				addr, err := excelize.CoordinatesToCellName(cindex+1, rindex+1)
				if err != nil {
					return err
				}
				x.SetCellValue(sheetName, addr, val)
			}
		} else {
			for i := len(columns); i < len(fields); i++ {
				colName, err := excelize.ColumnNumberToName(i + 1)
				if err != nil {
					return err
				}

				col := Column{Name: "#" + colName}
				if i := columnHints.FindByName(col.Name); i != -1 {
					col = columnHints[i]
				}
				columns = append(columns, col)
			}

			for cindex, val := range fields {
				g := columns[cindex]
				//log.Print(g.Name, val, g.Type)

				addr, err := excelize.CoordinatesToCellName(cindex+1, rindex+1)
				if err != nil {
					return fmt.Errorf("%v: %v\n", g.Name, err)
				}
				//log.Println(addr, g.Name, val)

				if len(val) == 0 {
					continue
				}

				if !c.GuessType {
					x.SetCellValue(sheetName, addr, val)

					continue
				}

				typ := ColType("")
				var fvalue float64
				var tvalue time.Time

				guessed := false
				switch g.Type {
				case TypeText:
					guessed = true
					typ = TypeText

				case TypeNumber:
					if f, err := strconv.ParseFloat(val, 64); err == nil {
						guessed = true
						typ = TypeNumber
						fvalue = f
					}

				case TypeDate:
					var ptns []string = translateDatePatterns(g.InputFormat)
					for _, ptn := range ptns {
						if t, err := time.Parse(ptn, val); err == nil {
							guessed = true
							typ = TypeDate
							tvalue = t
							break
						}
					}

				case TypeTime:
					var ptns []string = translateTimePatterns(g.InputFormat)
					for _, ptn := range ptns {
						if t, err := time.Parse(ptn, val); err == nil {
							guessed = true
							typ = TypeTime
							tvalue = t
							break
						}
					}

				case TypeDatetime:
					if t, err := time.Parse(g.InputFormat, val); err == nil {
						guessed = true
						typ = TypeDatetime
						tvalue = t
					}

				default: // nop
				}

				if !guessed && val[0] == '\'' {
					guessed = true
					typ = TypeText
				}
				if !guessed && val[0] == '0' {
					guessed = true
					typ = TypeText
				}
				if !guessed {
					if t, err := time.Parse(c.DatetimeFmt, val); err == nil {
						guessed = true
						typ = TypeDatetime
						tvalue = t
					}
				}
				if !guessed {
					for _, ptn := range datePtns {
						if len(ptn) != len(val) {
							continue
						}
						if t, err := time.Parse(ptn, val); err == nil {
							guessed = true
							typ = TypeDate
							tvalue = t
							break
						}
					}
				}
				if !guessed {
					for _, ptn := range timePtns {
						if len(ptn) != len(val) {
							continue
						}
						if t, err := time.Parse(ptn, val); err == nil {
							guessed = true
							typ = TypeTime
							tvalue = t
							break
						}
					}
				}
				if !guessed {
					if f, err := strconv.ParseFloat(val, 64); err == nil {
						guessed = true
						typ = TypeNumber
						fvalue = f
					}
				}

				//log.Println(addr, g.Name, val, typ)

				if typ == TypeText {
					x.SetCellValue(sheetName, addr, val)

				} else if typ == TypeDatetime {
					//log.Println(addr, tvalue)
					x.SetCellValue(sheetName, addr, tvalue)
					x.SetCellStyle(sheetName, addr, addr, datetimeStyle)

				} else if typ == TypeDate {
					//log.Println(addr, tvalue)
					x.SetCellValue(sheetName, addr, tvalue)
					x.SetCellStyle(sheetName, addr, addr, dateStyle)

				} else if typ == TypeTime {
					if y, m, d := tvalue.Date(); y == 0 && m == 1 && d == 1 {
						tvalue = time.Date(1900, 1, 1, tvalue.Hour(), tvalue.Minute(), tvalue.Second(), tvalue.Nanosecond(), tvalue.Location())
					}
					//log.Println(addr, tvalue)
					x.SetCellValue(sheetName, addr, tvalue)
					x.SetCellStyle(sheetName, addr, addr, timeStyle)

				} else if typ == TypeNumber {
					//log.Println(addr, fvalue)
					x.SetCellValue(sheetName, addr, fvalue)

				} else {
					x.SetCellValue(sheetName, addr, val)
				}

				columns[cindex] = g
			}
		}
		rindex++
	}

	x.SetActiveSheet(1)
	err = x.SaveAs(c.Output)
	if err != nil {
		return err
	}

	return nil
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

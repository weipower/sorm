package sorm

import (
	"github.com/Luxurioust/excelize"

	"github.com/bitly/go-simplejson"

	"bytes"
	"database/sql"
	"reflect"
	"strconv"
	"time"
)

//   __      __       .__
//  /  \    /  \ ____ |__|_____   ______  _  __ ___________    _________________  _____
//  \   \/\/   // __ \|  \____ \ /  _ \ \/ \/ // __ \_  __ \  /  ___/  _ \_  __ \/     \
//   \        /\  ___/|  |  |_> >  <_> )     /\  ___/|  | \/  \___ (  <_> )  | \/  Y Y  \
//    \__/\  /  \___  >__|   __/ \____/ \/\_/  \___  >__|    /____  >____/|__|  |__|_|  /
//         \/       \/   |__|                      \/             \/                  \/

const SHEETNAME = "Sheet1"

func Rows2Xls(rows *sql.Rows, filename string) error {

	f := excelize.NewFile()
	f.NewSheet(SHEETNAME)

	columns, err := rows.Columns()
	if err != nil {
		return err
	}
	for i, cellvalue := range columns {
		ord := int('A') + i
		cellname := string(rune(ord)) + "1"
		f.SetCellValue(SHEETNAME, cellname, cellvalue)
	}

	count := len(columns)

	values := make([]interface{}, count)
	valuePtrs := make([]interface{}, count)
	rindex := 1
	for rows.Next() {
		for i := 0; i < count; i++ {
			valuePtrs[i] = &values[i]
		}

		rows.Scan(valuePtrs...)

		for i, _ := range columns {
			var cellvalue interface{}
			val := values[i]
			var t time.Time

			if reflect.TypeOf(val) == reflect.TypeOf(t) {
				tt := val.(time.Time)

				var timestr string
				if tt.IsZero() {
					timestr = ""
				} else {
					tt2 := tt.Unix() - 8*3600
					timestr = time.Unix(tt2, 0).Format("2006-01-02 15:04:05")
				}
				cellvalue = timestr

			} else {
				b, ok := val.([]byte)
				if ok {
					cellvalue = string(b)
				} else {
					cellvalue = val
				}
			}

			ord := int('A') + i
			cellname := string(rune(ord)) + strconv.Itoa(rindex+1)

			f.SetCellValue(SHEETNAME, cellname, cellvalue)
		}
		rindex = rindex + 1
	}

	return f.SaveAs(filename)
}

func Json2xls(json, filename string) error {

	buf := bytes.NewBuffer([]byte(json))

	js, err := simplejson.NewFromReader(buf)
	if err != nil {
		return err
	}

	rows, err := js.Get("rows").Array()
	if err != nil {
		return err
	}

	f := excelize.NewFile()
	f.NewSheet(SHEETNAME)

	var cols []string
	for i, row := range rows {
		if each_map, ok := row.(map[string]interface{}); ok {

			if i == 0 {
				colindex := 0
				for k, _ := range each_map {
					ord := int('A') + colindex

					cellname := string(rune(ord)) + "1"

					f.SetCellValue(SHEETNAME, cellname, k)
					cols = append(cols, k)

					colindex = colindex + 1
				}
			}

			colindex := 0

			for _, v := range each_map {
				ord := int('A') + colindex
				cellname := string(rune(ord)) + strconv.Itoa(i+2)

				f.SetCellValue(SHEETNAME, cellname, v)

				colindex = colindex + 1
			}

		}
	}

	return f.SaveAs(filename)

}

package sorm

import (
	"fmt"
	"testing"

	_ "github.com/go-sql-driver/mysql"

	"database/sql"
)

func TestRows2xls(t *testing.T) {

	connstr := "root:1234@tcp(mysql.com:3306)/mysql?charset=utf8mb4&parseTime=true&loc=Local"

	db, err := sql.Open("mysql", connstr)
	if err != nil {
		panic(err)
	}

	defer db.Close()

	sql := `SELECT host as "Host", user as "User" FROM user`

	rows, err := db.Query(sql)
	if err != nil {
		panic(err)
	}
	defer rows.Close()

	if err = Rows2Xls(rows, "demo_one.xlsx"); err != nil {
		fmt.Println("can't save xls file.")
		panic(err)
	}

}

func TestJson2xls(t *testing.T) {
	json := `
	{"rows" :[
	{"a": 1,"b": "2","c": false,"d":null},
	{"a": 3,"b": "4","c": false,"d":null}]}`

	Json2xls(json, "demo_two.xlsx")
}

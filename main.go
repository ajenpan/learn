package main

import (
	"fmt"
	"log"
	"math/rand"
	"time"

	"github.com/xuri/excelize/v2"
)

func Shuffle(raw string) string {
	b := []byte(raw)
	rand.Seed(time.Now().UnixNano())
	rand.Shuffle(len(b), func(i, j int) {
		b[i], b[j] = b[j], b[i]
	})
	return string(b)
}

const jp50rol = "aiueo"
const jp50row = "akstnhmyrw"

var JP50pin = []string{
	"あいうえお",
	"かきくけこ",
	"さしすせそ",
	"たちつてと",
	"なにぬねの",
	"はひふへほ",
	"まみむめも",
	"や ゆ よ",
	"らりるれろ",
	"わ   を"}

var JP50pian = []string{
	"アイウエオ",
	"カキクケコ",
	"サシスセソ",
	"タチツテト",
	"ナニヌネノ",
	"ハヒフヘホ",
	"マミムメモ",
	"ヤ ユ ヨ",
	"ラリルレロ",
	"ワ   ヲ"}

func GenRandExcel() {
	f := excelize.NewFile()

	const sheetName = "Sheet1"
	cols := Shuffle(jp50rol)
	rows := Shuffle(jp50row)

	for i, row := range rows {
		pos := ExcelPos(i+2, 1)
		fmt.Println(pos)
		f.SetCellStr(sheetName, pos, string(row))
	}

	for i, col := range cols {
		pos := ExcelPos(1, i+2)
		fmt.Println(pos)
		f.SetCellStr(sheetName, pos, string(col))
	}

	if err := f.SaveAs(fmt.Sprintf("jp50_%s.xlsx", time.Now().Format("20060102150405"))); err != nil {
		log.Fatal(err)
	}
}

func CheckExcel() {
	filename := ""
	xlsx, err := excelize.OpenFile(filename)
	if err != nil {
		fmt.Println(err)
		return
	}

	xlsx.GetCellValue("Sheet1", "A1")
}

func main() {
	GenRandExcel()
}

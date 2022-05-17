package main

import "strconv"

func ExcelRow(i int) string {
	return strconv.Itoa(i)
}

func ExcelCol(i int) string {
	return string(rune(i - 1 + 'A'))
}

func ExcelPos(ri, ci int) string {
	return ExcelCol(ci) + ExcelRow(ri)
}

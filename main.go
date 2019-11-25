package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

func Headings(wrkbook *excelize.File, row string) {
	curcell := ""
	i := 0
	for i = 0; i < 2; i++ {
		curcell = row + strconv.Itoa(i+1)
		wrkbook.SetCellValue("Sheet1", curcell, 2.02)
	}
	curcell = row + strconv.Itoa(i+1)
	formula := TwoCellFormula("Sum", row+strconv.Itoa(1), row+strconv.Itoa(2))
	wrkbook.SetCellFormula("Sheet1", curcell, formula)
}

func TwoCellFormula(formulaType, firstCell, lastCell string) string {
	return formulaType + "(" + firstCell + ":" + lastCell + ")"
}

func SingleHistoricalFreq(freq float64, impact int) string {
	return ""
}

func main() {

	wrkbook := excelize.NewFile()
	wrkbook.NewSheet("Sheet1")
	Headings(wrkbook, "A")
	err := wrkbook.SaveAs("C:/Users/johnn/Documents/projects-yugioh/datasheets/blanksheet.xlsx")
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println("check sheet")
}

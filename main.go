package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
)

func SingleHistoricalFreq(freq float64, impact int) float64 {
	return 0.0
}

func main() {

	wrkbook := excelize.NewFile()
	wrkbook.NewSheet("Sheet1")
	wrkbook.SetCellValue("Sheet1", "A1", 2.02)
	wrkbook.SetCellValue("Sheet1", "A2", 1.01)
	wrkbook.SetCellFormula("Sheet1", "A3", "SUM(A1:A2)")
	err := wrkbook.SaveAs("C:/Users/johnn/Documents/projects-yugioh/datasheets/blanksheet.xlsx")
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println("check sheet")
}

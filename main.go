package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

var CardNames []string
var DeckNames []string

var lettermapping map[int]string

//TODO: CHANGE THIS TO WORK WITH ROW INSTEAD OF COLUMN
func Headings(wrkbook *excelize.File, column string) {
	//start from E, go to end of thing
	curcell := ""
	i := 0
	for i = 0; i < 2; i++ {
		curcell = column + strconv.Itoa(i+1)
		wrkbook.SetCellValue("Sheet1", curcell, 2.02)
	}
	curcell = column + strconv.Itoa(i+1)
	formula := TwoCellFormula("Sum", column+strconv.Itoa(1), column+strconv.Itoa(2))
	wrkbook.SetCellFormula("Sheet1", curcell, formula)
}

func Matchups(wrkbook *excelize.File, row int) {
	//set up matchups in B column
}

func ScoreVMean(wrkbook *excelize.File, row int) {
	//set up score vs mean in last column before moving on to next section
}

func TwoCellFormula(formulaType, firstCell, lastCell string) string {
	return formulaType + "(" + firstCell + ":" + lastCell + ")"
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

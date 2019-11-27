package main

import (
	"bufio"
	"fmt"
	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var CardNames []string
var DeckNames []string

var lettermapping = map[int]string{
	0:  "A",
	1:  "B",
	2:  "C",
	3:  "D",
	4:  "E",
	5:  "F",
	6:  "G",
	7:  "H",
	8:  "I",
	9:  "J",
	10: "K",
	11: "L",
	12: "M",
	13: "N",
	14: "O",
	15: "P",
	16: "Q",
	17: "R",
	18: "S",
	19: "T",
	20: "U",
	21: "V",
	22: "W",
	23: "X",
	24: "Y",
	25: "Z",
}

//---------------------------------------------------------------------------------
// SETUP FUNCTIONS
//---------------------------------------------------------------------------------

/**
* Setup Raw Score section
* Params: *excelize.File
* Returns: Row that it ends on
 */
func RawScoreSection(wrkbook *excelize.File) {

}

//---------------------------------------------------------------------------------
// HELPER FUNCTIONS
//---------------------------------------------------------------------------------

func Headings(wrkbook *excelize.File, row int, cursheet string) {
	//start from E, go to end of thing
	curcell := ""
	i := 0
	for i = 0; i < len(CardNames); i++ {
		curcell = lettermapping[4+i] + strconv.Itoa(row)
		wrkbook.SetCellValue(cursheet, curcell, CardNames[i])
	}
}

func Matchups(wrkbook *excelize.File, row int, cursheet string) {
	//set up matchups in B column
	wrkbook.SetCellValue(cursheet, "B"+strconv.Itoa(row), "Matchup")
	row += 1
	index := row
	for index < row+len(DeckNames) {
		wrkbook.SetCellValue(cursheet, "B"+strconv.Itoa(index), DeckNames[index-row])
		index++
	}
}

func SetStandardFormulas(wrkbook *excelize.File, row int, cursheet string) {
	//separate counter for looping through formula
	index := 3
	//first loop: row
	for i := row; i < len(DeckNames)+row; i++ {
		//second loop: column
		for j := 4; j < len(CardNames)+4; j++ {
			curCell := lettermapping[j] + strconv.Itoa(i)
			formula := "D" + strconv.Itoa(i) + "*" + lettermapping[j] + strconv.Itoa(index)
			wrkbook.SetCellFormula(cursheet, curCell, formula)
		}
		index++
	}
}

func WeightedSum(wrkbook *excelize.File, row int, space bool, cursheet string) {
	thisrow := 0
	if !space {
		thisrow = row + len(DeckNames) + 1
	} else {
		thisrow = row + len(DeckNames) + 2
	}
	wrkbook.SetCellValue(cursheet, "B"+strconv.Itoa(thisrow), "Weighted Sum")
	for j := 4; j < len(CardNames)+4; j++ {
		curCell := lettermapping[j] + strconv.Itoa(thisrow)
		formula := "SUM(" + lettermapping[j] + strconv.Itoa(row+1) + ":"
		if !space {
			formula = formula + lettermapping[j] + strconv.Itoa(thisrow-1) + ")"
		} else {
			formula = formula + lettermapping[j] + strconv.Itoa(thisrow-2) + ")"
		}
		wrkbook.SetCellFormula(cursheet, curCell, formula)
	}
}

func ScoreVsMean(wrkbook *excelize.File, row int, space bool, cursheet string) {
	//recalculate row to correct place
	if space {
		row += len(DeckNames) + 4
	} else {
		row += len(DeckNames) + 3
	}
	//set name of row
	wrkbook.SetCellValue(cursheet, "B"+strconv.Itoa(row), "Score Vs Mean")
	//set row string
	formstr := strconv.Itoa(row - 1)
	//grab weighted sums and conver to formula
	formula := "SUM(E" + formstr + ":" + lettermapping[len(CardNames)+3] + formstr + ")/" + strconv.Itoa(len(CardNames))
	//save formula cell to save work later
	formcell := "C" + strconv.Itoa(row)
	wrkbook.SetCellFormula(cursheet, formcell, formula)
	//loop through and set score vs mean formula
	for j := 4; j < len(CardNames)+4; j++ {
		formula = lettermapping[j] + formstr + "/" + formcell
		wrkbook.SetCellFormula(cursheet, lettermapping[j]+strconv.Itoa(row), formula)
	}

}

func Frequency(wrkbook *excelize.File, row int, cursheet string) {
	//set Total Number of Decks
	thisrow := row + len(DeckNames) + 1
	formula := "SUM(C" + strconv.Itoa(row+1) + ":C" + strconv.Itoa(thisrow-1) + ")"
	wrkbook.SetCellFormula(cursheet, "C"+strconv.Itoa(thisrow), formula)
	//set new frequency formulas
	row += 1
	curCell := ""
	for row < thisrow {
		curCell = "D" + strconv.Itoa(row)
		formula = "C" + strconv.Itoa(row) + "/C" + strconv.Itoa(thisrow)
		wrkbook.SetCellFormula(cursheet, curCell, formula)
		row++
	}
}

func main() {

	//open defensive cards text file
	file, err := os.Open("C:/Users/johnn/Documents/projects-yugioh/datasheets/defensivecards.txt")

	if err != nil {
		panic(err)
	}

	scanner := bufio.NewScanner(file)
	scanner.Split(bufio.ScanLines)

	//initialize CardNames slice

	for scanner.Scan() {
		if CardNames == nil {
			CardNames = []string{scanner.Text()}
		} else {
			CardNames = append(CardNames, scanner.Text())
		}
	}

	file.Close()

	//open decks text file
	file, err = os.Open("C:/Users/johnn/Documents/projects-yugioh/datasheets/metadecks.txt")

	if err != nil {
		panic(err)
	}

	scanner = bufio.NewScanner(file)
	scanner.Split(bufio.ScanLines)

	//initialize CardNames slice

	for scanner.Scan() {
		if DeckNames == nil {
			DeckNames = []string{scanner.Text()}
		} else {
			DeckNames = append(DeckNames, scanner.Text())
		}
	}

	file.Close()

	wrkbook := excelize.NewFile()
	cursheet := "Sheet1"
	wrkbook.NewSheet(cursheet)
	index := 2
	Headings(wrkbook, index, cursheet)
	Matchups(wrkbook, index, cursheet)
	WeightedSum(wrkbook, index, false, cursheet)
	ScoreVsMean(wrkbook, index-3, false, cursheet)
	index += len(DeckNames) + 5
	Headings(wrkbook, index, cursheet)
	Matchups(wrkbook, index, cursheet)
	SetStandardFormulas(wrkbook, index+1, cursheet)
	WeightedSum(wrkbook, index, true, cursheet)
	Frequency(wrkbook, index, cursheet)
	ScoreVsMean(wrkbook, index-3, true, cursheet)
	index += len(DeckNames) + 5

	err = wrkbook.SaveAs("C:/Users/johnn/Documents/projects-yugioh/datasheets/blanksheet.xlsx")
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println("check sheet")
}

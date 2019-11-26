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
}

func Matchups(wrkbook *excelize.File, row int) {
	//set up matchups in B column
	wrkbook.SetCellValue("Sheet1", "B"+strconv.Itoa(row), "Matchup")
	row += 1
	index := row
	for index < row+len(DeckNames) {
		wrkbook.SetCellValue("Sheet1", "B"+strconv.Itoa(index), DeckNames[index-row])
		index++
	}
}

func ScoreVMean(wrkbook *excelize.File, row int) {
	//set up score vs mean in last column before moving on to next section
}

func TwoCellFormula(formulaType, firstCell, lastCell string) string {
	return formulaType + "(" + firstCell + ":" + lastCell + ")"
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
	wrkbook.NewSheet("Sheet1")
	Matchups(wrkbook, 10)
	err = wrkbook.SaveAs("C:/Users/johnn/Documents/projects-yugioh/datasheets/blanksheet.xlsx")
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println("check sheet")
}

package main

import (
	"bufio"
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"strconv"
	"time"

	"github.com/tealeg/xlsx"
)

var MAX_ROW = 1008500

type paraDataForXlsx struct {
	fp   *xlsx.File
	cNum int
}

var column = []string{
	"policyID", "statecode", "county", "eq_site_limit",
	"hu_site_limit", "fl_site_limit", "fr_site_limit",
	"tiv_2011", "tiv_2012", "eq_site_deductible", "hu_site_deductible",
	"fl_site_deductible", "fr_site_deductible", "point_latitude",
	"point_longitude", "line", "construction", "point_granularity",
}

func inputRow(cutNum, currentRow, sheetNum int, xlsxFile *xlsx.File, rows [][]string) paraDataForXlsx {

	sheetName := "Sheet" + strconv.Itoa(sheetNum)
	iRows := rows
	fmt.Println("Dealing with Sheet Number ", sheetNum)
	xlsheet, err := xlsxFile.AddSheet(sheetName)

	if err != nil {
		fmt.Printf(err.Error())
	}

	for i, row := range iRows {

		fmt.Print("\033[G\033[K") //restore the cursor position and clear the line
		fmt.Printf("Putting %d row inside Sheet Number %d \n", cutNum, sheetNum)

		xlrow := xlsheet.AddRow()
		for _, value := range row {
			xlcell := xlrow.AddCell()
			xlcell.Value = value

		}
		if cutNum >= (MAX_ROW * sheetNum) {
			currentRow = i
			break
		}
		cutNum++
		fmt.Print("\033[A") // move the cursor up
	}

	fmt.Println()

	if cutNum >= (MAX_ROW * sheetNum) {
		sheetNum++
		fmt.Printf("Put %d row inside the Sheet Number %d\n", cutNum, sheetNum)
		fmt.Println("Making a New Sheet")
		return inputRow(cutNum, currentRow, sheetNum, xlsxFile, iRows)
	}

	fmt.Println("Returning the file pointer")
	tmpParaData := new(paraDataForXlsx)
	tmpParaData.cNum = cutNum
	tmpParaData.fp = xlsxFile
	return *tmpParaData

}

func csv2xlsx(csvFile, xlsxFile string) {

	file, err := os.Open(csvFile)

	if err != nil {
		log.Panic(err)
	}

	defer file.Close()

	// rdr := csv.NewReader(bufio.NewReader(file))
	r := csv.NewReader(bufio.NewReader(file))
	rows, err := r.ReadAll()

	// row 읽어 들이기
	if err != nil {
		log.Panic(err)
	}

	cutNum := 0
	currentRow := 0

	sheetTime := time.Now()

	fmt.Println("Printing File It will take a lot of time...")

	xlsxFileIns := xlsx.NewFile()

	tmpParaData := inputRow(cutNum, currentRow, 1, xlsxFileIns, rows)

	fmt.Println("Saving the file into xlsx file...")

	err = tmpParaData.fp.Save(xlsxFile)

	if err != nil {
		fmt.Println(err.Error())
	}

	fmt.Println("Saving has been completed!")
	endTime := time.Since(sheetTime)

	fmt.Println("==>Finished and it took ", endTime)
	fmt.Println("==>Row count :", tmpParaData.cNum)

}

func main() {

	args := os.Args[1:]

	csvFile := args[0]
	xlsxFile := args[1]
	csv2xlsx(csvFile, xlsxFile)

}

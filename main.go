package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	fmt.Println("Hello World!")
	myFile := excelize.NewFile()
	//Manually Set the cell values
	myFile.SetCellValue("Sheet1", "A1", "Heading 1")
	myFile.SetCellValue("Sheet1", "A2", "I'm in cell A2!")
	myFile.SetCellValue("Sheet1", "A3", "I'm in cell A3!")

	myFile.SetCellValue("Sheet1", "B1", "Heading 2")
	myFile.SetCellValue("Sheet1", "B2", "I'm in cell B2!")
	myFile.SetCellValue("Sheet1", "B3", "I'm in cell B3!")
	
	myFile.SetCellValue("Sheet1", "C1", "Heading 3")
	myFile.SetCellValue("Sheet1", "C2", "I'm in cell C2!")
	myFile.SetCellValue("Sheet1", "C3", "I'm in cell C3!")

	//read from a single cell
	columnA, err := myFile.GetCellValue("Sheet1", "A1")
	fmt.Println(columnA, err)

	//Read entire row
	rows, err := myFile.GetRows("Sheet1")
	//dump all rows
	fmt.Println(rows, err)
	//dump first value of rows array
	fmt.Println(rows[1], err)
	fmt.Println(rows[2], err)

	fmt.Println(rows[1:][0][1])
	fmt.Println(rows[1:][1][1])

	//slurp up contents of column, excluding the header column:
	fmt.Println("Starting for loop....")
	for _, v := range rows[1:] {
		fmt.Println(v[1])
	}

	//loop over rows
	// for _, row := range rows {
	// 	fmt.Println("Value of row is: ", row)
	// 	for _, colCell := range row {
	// 		fmt.Println("The value of colCell is: ", colCell)
	// 	}
	// 	fmt.Println()
	// }

	myFile.SaveAs("myworkbook.xlsx")

}
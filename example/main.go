package main

import (
	"fmt"

	. "github.com/darmawan01/xlsx_utilities"
)

func main() {
	type Person struct {
		Name string
		Age  int
	}

	// Create sample data
	data := []Person{
		{Name: "Alice", Age: 30},
		{Name: "Bob", Age: 25},
		{Name: "Charlie", Age: 35},
	}

	// Convert struct to ExcelData
	excelData, err := FromStruct(data)
	if err != nil {
		fmt.Println("Error converting struct to ExcelData:", err)
		return
	}

	// Write to Excel file
	err = excelData.ToExcel("people.xlsx")
	if err != nil {
		fmt.Println("Error writing Excel file:", err)
		return
	}

	// Read from Excel file
	readExcelData, err := FromExcel[Person]("people.xlsx")
	if err != nil {
		fmt.Println("Error reading Excel file:", err)
		return
	}

	// Convert ExcelData back to struct
	result := readExcelData.ToStruct()

	fmt.Printf("Imported %d records\n", len(result.Data))

	if len(result.Errors) > 0 {
		fmt.Println("Import errors occurred:")
		fmt.Println(FormatImportErrors(result.Errors))
	} else {
		fmt.Println("No import errors")
	}

	// Use the imported data
	for _, person := range result.Data {
		fmt.Printf("%+v\n", person)
	}
}

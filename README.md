# Excel Utility Package for Go

This package provides a set of utilities for working with Excel files in Go, allowing easy conversion between Go structs and Excel spreadsheets.

## Features

- Convert slices of structs to Excel files
- Read Excel files into slices of structs
- Generic implementation for flexibility with different struct types
- Error handling for import/export operations
- Type-safe operations with Go generics

## Installation

To use this package, you need to have Go installed (version 1.18+ for generics support). Then, you can install the package using:

```bash
go get github.com/darmawan01/xlsx_utilities
```

Also, make sure to install the required dependencies:

```bash
go get github.com/xuri/excelize/v2
go get github.com/stretchr/testify/assert # for running tests
```

## Usage

Here's a quick example of how to use the main features of the package:

```go
package main

import (
    "fmt"
    "github.com/darmawan01/xlsx_utilities"
)

type Person struct {
    Name string
    Age  int
}

func main() {
    // Sample data
    data := []Person{
        {Name: "Alice", Age: 30},
        {Name: "Bob", Age: 25},
        {Name: "Charlie", Age: 35},
    }

    // Convert struct to ExcelData and write to file
    excelData, err := xlsx_utilities.FromStruct(data)
    if err != nil {
        fmt.Println("Error converting struct to ExcelData:", err)
        return
    }

    err = excelData.ToExcel("people.xlsx")
    if err != nil {
        fmt.Println("Error writing Excel file:", err)
        return
    }

    // Read from Excel file
    readExcelData, err := xlsx_utilities.FromExcel[Person]("people.xlsx")
    if err != nil {
        fmt.Println("Error reading Excel file:", err)
        return
    }

    // Convert ExcelData back to struct
    result := readExcelData.ToStruct()

    fmt.Printf("Imported %d records\n", len(result.Data))

    if len(result.Errors) > 0 {
        fmt.Println("Import errors occurred:")
        fmt.Println(xlsx_utilities.FormatImportErrors(result.Errors))
    } else {
        fmt.Println("No import errors")
    }

    // Use the imported data
    for _, person := range result.Data {
        fmt.Printf("%+v\n", person)
    }
}
```

## API Reference

### Types

- `ExcelData[T comparable]`: Represents Excel data for a given struct type T.
- `ImportResult[T comparable]`: Represents the result of importing Excel data to a struct, including any errors.
- `ImportError`: Represents an error that occurred during the import process.

### Functions

- `NewExcelData[T comparable](headers []string) *ExcelData[T]`: Creates a new ExcelData instance.
- `FromStruct[T comparable](data []T) (*ExcelData[T], error)`: Converts a slice of structs to ExcelData.
- `FromExcel[T comparable](filename string) (*ExcelData[T], error)`: Reads an Excel file into ExcelData.
- `FormatImportErrors(errors []ImportError) string`: Formats import errors into a readable string.

### Methods

- `(ed *ExcelData[T]) AddRow(row []interface{}) error`: Adds a new row to the ExcelData.
- `(ed *ExcelData[T]) ToExcel(filename string) error`: Generates an Excel file from the ExcelData.
- `(ed *ExcelData[T]) ToStruct() ImportResult[T]`: Converts ExcelData to a slice of struct T and collects import errors.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
# Excel Utility Package for Go

This package provides a set of utilities for working with Excel files in Go, allowing easy conversion between Go structs (including nested structs and custom types) and Excel spreadsheets.

## Features

- Convert slices of structs (including nested structs and custom types) to Excel files
- Read Excel files into slices of structs
- Generic implementation for flexibility with different struct types
- Error handling for import/export operations
- Type-safe operations with Go generics
- Support for nested structs with flattened headers
- Custom type handling with user-definable converters and parsers

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

Here's a quick example of how to use the main features of the package, including nested structs and custom types:

```go
package main

import (
    "fmt"
    "time"
    "reflect"
    "github.com/darmawan01/xlsx_utilities"
)

type Address struct {
    Street string
    City   string
}

type CustomID int

type Person struct {
    Name      string
    Age       int
    Address   Address
    BirthDate time.Time
    ID        CustomID
}

func (c CustomID) String() string {
    return fmt.Sprintf("ID-%05d", int(c))
}

func main() {
    // Register custom type handlers
    xlsx_utilities.RegisterTypeConverter(reflect.TypeOf(CustomID(0)), func(i interface{}) (string, error) {
        id, ok := i.(CustomID)
        if !ok {
            return "", fmt.Errorf("expected CustomID, got %T", i)
        }
        return id.String(), nil
    })

    xlsx_utilities.RegisterTypeParser(reflect.TypeOf(CustomID(0)), func(s string) (interface{}, error) {
        var id int
        _, err := fmt.Sscanf(s, "ID-%05d", &id)
        if err != nil {
            return nil, err
        }
        return CustomID(id), nil
    })

    // Sample data with nested structs and custom types
    data := []Person{
        {
            Name: "Alice",
            Age:  30,
            Address: Address{
                Street: "123 Main St",
                City:   "New York",
            },
            BirthDate: time.Date(1990, 1, 1, 0, 0, 0, 0, time.UTC),
            ID:        CustomID(1),
        },
        {
            Name: "Bob",
            Age:  25,
            Address: Address{
                Street: "456 Elm St",
                City:   "San Francisco",
            },
            BirthDate: time.Date(1995, 5, 15, 0, 0, 0, 0, time.UTC),
            ID:        CustomID(2),
        },
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
- `CustomTypeConverter`: Function type for custom type conversions.
- `CustomTypeParser`: Function type for parsing custom types from strings.

### Functions

- `NewExcelData[T comparable](headers []string) *ExcelData[T]`: Creates a new ExcelData instance.
- `FromStruct[T comparable](data []T) (*ExcelData[T], error)`: Converts a slice of structs (including nested structs and custom types) to ExcelData.
- `FromExcel[T comparable](filename string) (*ExcelData[T], error)`: Reads an Excel file into ExcelData.
- `FormatImportErrors(errors []ImportError) string`: Formats import errors into a readable string.
- `RegisterTypeConverter(t reflect.Type, converter CustomTypeConverter)`: Registers a custom type converter.
- `RegisterTypeParser(t reflect.Type, parser CustomTypeParser)`: Registers a custom type parser.

### Methods

- `(ed *ExcelData[T]) AddRow(row []interface{}) error`: Adds a new row to the ExcelData.
- `(ed *ExcelData[T]) ToExcel(filename string) error`: Generates an Excel file from the ExcelData.
- `(ed *ExcelData[T]) Save(filename string) error`: Saves the Excel file.
- `(ed *ExcelData[T]) ToFile() *excelize.File`: Generates an Excel file from the ExcelData and returns the file object.
- `(ed *ExcelData[T]) ToStruct() ImportResult[T]`: Converts ExcelData to a slice of struct T and collects import errors.

## Nested Struct Support

The package supports nested structs when converting to and from Excel files. Headers for nested fields are flattened using space notation (e.g., "Address Street", "Address City").

## Custom Type Handling

The package now supports custom type handling through user-definable converters and parsers. Users can register custom type handlers for any type they need to work with in their Excel conversions. This allows for seamless integration of complex or domain-specific types in your Excel operations.

To use custom type handling:

1. Define your custom type and its string representation method (if needed).
2. Register a custom type converter using `RegisterTypeConverter`.
3. Register a custom type parser using `RegisterTypeParser`.

The package will then automatically use these handlers when converting to and from Excel.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
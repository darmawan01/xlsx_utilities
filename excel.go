package xlsx_utilities

import (
	"fmt"
	"os"
	"reflect"
	"strings"

	"github.com/xuri/excelize/v2"
)

// ExcelData represents a generic struct for Excel data
type ExcelData[T comparable] struct {
	Headers []string
	Rows    [][]interface{}
}

// ImportError represents an error that occurred during the import process
type ImportError struct {
	RowIndex int
	Header   string
	Value    interface{}
	Type     reflect.Type
	Err      error
}

// ImportResult represents the result of importing Excel data to a struct
type ImportResult[T comparable] struct {
	Data   []T
	Errors []ImportError
}

// Error returns a string representation of the ImportError
func (e ImportError) Error() string {
	return fmt.Sprintf("Row %d, Column '%s': cannot convert '%v' to type %v: %v", e.RowIndex, e.Header, e.Value, e.Type, e.Err)
}

// NewExcelData creates a new ExcelData instance
func NewExcelData[T comparable](headers []string) *ExcelData[T] {
	return &ExcelData[T]{
		Headers: headers,
		Rows:    make([][]interface{}, 0),
	}
}

// AddRow adds a new row to the ExcelData
func (ed *ExcelData[T]) AddRow(row []interface{}) error {
	if len(row) != len(ed.Headers) {
		return fmt.Errorf("row length (%d) does not match headers length (%d)", len(row), len(ed.Headers))
	}
	ed.Rows = append(ed.Rows, row)
	return nil
}

// ToExcel generates an Excel file from the ExcelData
func (ed *ExcelData[T]) ToExcel(filename string) error {
	return ed.Save(filename)
}

// Save the Excel file
func (ed *ExcelData[T]) Save(filename string) error {
	f := ed.ToFile()
	defer f.Close()

	return f.SaveAs(filename)
}

// ToFile generates an Excel file from the ExcelData
func (ed *ExcelData[T]) ToFile() *excelize.File {
	f := excelize.NewFile()

	// Write headers
	for col, header := range ed.Headers {
		cell := fmt.Sprintf("%s1", intToExcelColumn(col))
		f.SetCellValue("Sheet1", cell, header)
	}

	// Write data
	for rowIndex, row := range ed.Rows {
		for col, value := range row {
			cell := fmt.Sprintf("%s%d", intToExcelColumn(col), rowIndex+2)
			f.SetCellValue("Sheet1", cell, value)
		}
	}

	return f
}

// intToExcelColumn converts a 0-based column index to an Excel column name (A, B, C, ..., Z, AA, AB, etc.)
func intToExcelColumn(n int) string {
	result := ""
	for n >= 0 {
		result = string(rune('A'+n%26)) + result
		n = n/26 - 1
	}
	return result
}

// excelColumnToInt converts an Excel column name to a 0-based column index
func excelColumnToInt(columnName string) int {
	result := 0
	for _, char := range columnName {
		result = result*26 + int(char-'A') + 1
	}
	return result - 1
}

// FromExcel reads an Excel file into ExcelData
func FromExcel[T comparable](filename string) (*ExcelData[T], error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	f, err := excelize.OpenReader(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		return nil, err
	}

	if len(rows) < 2 {
		return nil, fmt.Errorf("excel file is empty or has no data rows")
	}

	ed := NewExcelData[T](rows[0])

	for _, row := range rows[1:] {
		interfaceRow := make([]interface{}, len(row))
		for i, cell := range row {
			interfaceRow[i] = convertCellValue(cell)
		}
		ed.Rows = append(ed.Rows, interfaceRow)
	}

	return ed, nil
}

// ToStruct converts ExcelData to a slice of struct T and collects import errors
func (ed *ExcelData[T]) ToStruct() ImportResult[T] {
	var result []T
	var importErrors []ImportError

	t := reflect.TypeOf((*T)(nil)).Elem()

	for rowIndex, row := range ed.Rows {
		item := reflect.New(t).Elem()
		rowErrors := []ImportError{}

		for i, header := range ed.Headers {
			if i < len(row) {
				err := setNestedField(item, header, row[i])
				if err != nil {
					rowErrors = append(rowErrors, ImportError{
						RowIndex: rowIndex + 2, // +2 because Excel rows are 1-indexed and we skip the header
						Header:   header,
						Value:    row[i],
						Type:     reflect.TypeOf(row[i]),
						Err:      err,
					})
				}
			}
		}

		if len(rowErrors) == 0 {
			result = append(result, item.Interface().(T))
		}
		importErrors = append(importErrors, rowErrors...)
	}

	return ImportResult[T]{
		Data:   result,
		Errors: importErrors,
	}
}

// FromStruct converts a slice of struct T to ExcelData, supporting nested structs
func FromStruct[T comparable](data []T) (*ExcelData[T], error) {
	if len(data) == 0 {
		return nil, fmt.Errorf("input slice is empty")
	}

	t := reflect.TypeOf((*T)(nil)).Elem()
	headers, err := getStructHeaders(t)
	if err != nil {
		return nil, fmt.Errorf("error getting headers: %v", err)
	}

	ed := NewExcelData[T](headers)

	for i, item := range data {
		row, err := getStructValues(reflect.ValueOf(item))
		if err != nil {
			return nil, fmt.Errorf("error getting values for item %d: %v", i, err)
		}

		if len(row) != len(headers) {
			return nil, fmt.Errorf("mismatch between headers (%d) and values (%d) for item %d", len(headers), len(row), i)
		}

		err = ed.AddRow(row)
		if err != nil {
			return nil, fmt.Errorf("error adding row %d: %v", i, err)
		}
	}

	return ed, nil
}

// FormatImportErrors returns a formatted string of all import errors
func FormatImportErrors(errors []ImportError) string {
	var sb strings.Builder
	for _, err := range errors {
		sb.WriteString(err.Error())
		sb.WriteString("\n")
	}
	return sb.String()
}

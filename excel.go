package xlsx_utilities

import (
	"fmt"
	"reflect"
	"strings"
	"time"

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
		cell := fmt.Sprintf("%c1", 'A'+col)
		f.SetCellValue("Sheet1", cell, header)
	}

	// Write data
	for rowIndex, row := range ed.Rows {
		for col, value := range row {
			cell := fmt.Sprintf("%c%d", 'A'+col, rowIndex+2)
			f.SetCellValue("Sheet1", cell, value)
		}
	}

	return f
}

// FromExcel reads an Excel file into ExcelData
func FromExcel[T comparable](filename string) (*ExcelData[T], error) {
	f, err := excelize.OpenFile(filename)
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

		importErrors = append(importErrors, rowErrors...)
		if len(rowErrors) == 0 {
			result = append(result, item.Interface().(T))
		}
	}

	return ImportResult[T]{
		Data:   result,
		Errors: importErrors,
	}
}

// setNestedField sets the value of a potentially nested struct field
func setNestedField(v reflect.Value, fieldPath string, value interface{}) error {
	fields := strings.Split(fieldPath, " ")
	for i, field := range fields {
		if v.Kind() == reflect.Ptr {
			if v.IsNil() {
				v.Set(reflect.New(v.Type().Elem()))
			}
			v = v.Elem()
		}

		if v.Kind() != reflect.Struct {
			return fmt.Errorf("not a struct")
		}

		f := v.FieldByName(field)
		if !f.IsValid() {
			return fmt.Errorf("no such field: %s in obj", field)
		}

		if i == len(fields)-1 {
			return setField(f, value)
		}
		v = f
	}
	return nil
}

// FromStruct converts a slice of struct T to ExcelData, supporting nested structs
func FromStruct[T comparable](data []T) (*ExcelData[T], error) {
	if len(data) == 0 {
		return nil, fmt.Errorf("input slice is empty")
	}

	headers, err := getStructHeaders(reflect.TypeOf(data[0]))
	if err != nil {
		return nil, err
	}

	ed := NewExcelData[T](headers)

	for _, item := range data {
		row, err := getStructValues(reflect.ValueOf(item))
		if err != nil {
			return nil, err
		}
		ed.AddRow(row)
	}

	return ed, nil
}

// getStructHeaders returns a flattened list of headers for a struct, including nested structs
func getStructHeaders(t reflect.Type) ([]string, error) {
	var headers []string

	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)

		// Skip unexported fields
		if !field.IsExported() {
			continue
		}

		fieldType := field.Type
		// If it's a pointer, get the element type
		if fieldType.Kind() == reflect.Ptr {
			fieldType = fieldType.Elem()
		}

		if fieldType.Kind() == reflect.Struct && fieldType != reflect.TypeOf(time.Time{}) {
			nestedHeaders, err := getStructHeaders(fieldType)
			if err != nil {
				return nil, err
			}
			for _, header := range nestedHeaders {
				headers = append(headers, field.Name+" "+header)
			}
		} else {
			headers = append(headers, field.Name)
		}
	}

	return headers, nil
}

// getStructValues returns a flattened list of values for a struct, including nested structs and custom types
func getStructValues(v reflect.Value) ([]interface{}, error) {
	var values []interface{}

	for i := 0; i < v.NumField(); i++ {
		field := v.Field(i)
		fieldType := v.Type().Field(i)

		// Skip unexported fields
		if !fieldType.IsExported() {
			continue
		}

		// Dereference pointer if needed
		if field.Kind() == reflect.Ptr {
			if field.IsNil() {
				values = append(values, nil)
				continue
			}
			field = field.Elem()
		}

		if field.Kind() == reflect.Struct && field.Type() != reflect.TypeOf(time.Time{}) {
			nestedValues, err := getStructValues(field)
			if err != nil {
				return nil, err
			}
			values = append(values, nestedValues...)
		} else {
			value := field.Interface()
			if converter, ok := TypeConverters[field.Type()]; ok {
				convertedValue, err := converter(value)
				if err != nil {
					return nil, fmt.Errorf("error converting custom type: %v", err)
				}
				value = convertedValue
			}
			values = append(values, value)
		}
	}

	return values, nil
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

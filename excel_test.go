package xlsx_utilities

import (
	"fmt"
	"os"
	"reflect"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

type person struct {
	Name string
	Age  int
}

func TestExcelDataOperations(t *testing.T) {
	// Test data
	data := []person{
		{Name: "Alice", Age: 30},
		{Name: "Bob", Age: 25},
		{Name: "Charlie", Age: 35},
	}

	t.Run("FromStruct", func(t *testing.T) {
		excelData, err := FromStruct(data)
		assert.NoError(t, err)
		assert.Equal(t, []string{"Name", "Age"}, excelData.Headers)
		assert.Len(t, excelData.Rows, 3)
		assert.Equal(t, []interface{}{"Alice", 30}, excelData.Rows[0])
	})

	t.Run("ToExcel", func(t *testing.T) {
		excelData, _ := FromStruct(data)
		err := excelData.ToExcel("test_people.xlsx")
		assert.NoError(t, err)

		// Verify file contents
		f, err := excelize.OpenFile("test_people.xlsx")
		assert.NoError(t, err)
		defer f.Close()

		cells, err := f.GetRows("Sheet1")
		assert.NoError(t, err)
		assert.Len(t, cells, 4) // Header + 3 data rows
		assert.Equal(t, []string{"Name", "Age"}, cells[0])
		assert.Equal(t, []string{"Alice", "30"}, cells[1])
	})

	t.Run("FromExcel", func(t *testing.T) {
		excelData, err := FromExcel[person]("test_people.xlsx")
		assert.NoError(t, err)
		assert.Equal(t, []string{"Name", "Age"}, excelData.Headers)
		assert.Len(t, excelData.Rows, 3)
		assert.Equal(t, []interface{}{"Alice", 30}, excelData.Rows[0])
	})

	t.Run("ToStruct", func(t *testing.T) {
		excelData, _ := FromExcel[person]("test_people.xlsx")
		result := excelData.ToStruct()
		assert.Len(t, result.Errors, 0)
		assert.Len(t, result.Data, 3)
		assert.Equal(t, person{Name: "Alice", Age: 30}, result.Data[0])

		// Check if returned data is of type []person
		_, ok := interface{}(result.Data).([]person)
		assert.True(t, ok, "result.Data should be of type []person")

		// Additional type check using reflection
		actualType := reflect.TypeOf(result.Data)
		expectedType := reflect.TypeOf([]person{})
		assert.Equal(t, expectedType, actualType, "result.Data should have type []person")

		// Check individual elements
		for _, item := range result.Data {
			_, ok := interface{}(item).(person)
			assert.True(t, ok, "Each item in result.Data should be of type person")
		}
	})

	t.Run("ToStruct with errors", func(t *testing.T) {
		// Create a new Excel file with an error
		f := excelize.NewFile()
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "B1", "Age")
		f.SetCellValue("Sheet1", "A2", "David")
		f.SetCellValue("Sheet1", "B2", "not_a_number")
		f.SaveAs("test_error.xlsx")
		defer os.Remove("test_error.xlsx")

		excelData, _ := FromExcel[person]("test_error.xlsx")
		result := excelData.ToStruct()
		assert.Len(t, result.Errors, 1)
		assert.Contains(t, result.Errors[0].Error(), "Row 2, Column 'Age': cannot convert 'not_a_number' to type int")
		assert.Len(t, result.Data, 0) // No valid data due to error

		// Check if returned data is of type []person even when empty
		_, ok := interface{}(result.Data).([]person)
		assert.True(t, ok, "result.Data should be of type []person even when empty")
	})

	os.Remove("test_people.xlsx")
}

func TestConvertCellValue(t *testing.T) {
	tests := []struct {
		input    string
		expected interface{}
	}{
		{"42", 42},
		{"3.14", 3.14},
		{"true", true},
		{"false", false},
		{"hello", "hello"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			result := convertCellValue(tt.input)
			assert.Equal(t, tt.expected, result)
		})
	}
}

func TestFormatImportErrors(t *testing.T) {
	errors := []ImportError{
		{RowIndex: 2, Header: "Age", Value: "not_a_number", Type: reflect.TypeOf(0), Err: fmt.Errorf("strconv.ParseInt: parsing \"not_a_number\": invalid syntax")},
		{RowIndex: 3, Header: "Name", Value: 42, Type: reflect.TypeOf(""), Err: fmt.Errorf("cannot convert int to string")},
	}

	formatted := FormatImportErrors(errors)
	assert.Contains(t, formatted, "Row 2, Column 'Age': cannot convert 'not_a_number' to type int")
	assert.Contains(t, formatted, "Row 3, Column 'Name': cannot convert '42' to type string")
}

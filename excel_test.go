package xlsx_utilities

import (
	"fmt"
	"os"
	"path/filepath"
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
	t.Run("NewExcelData", func(t *testing.T) {
		headers := []string{"Full Name", "Current Age"}
		excelData := NewExcelData[person](headers)

		assert.Equal(t, headers, excelData.Headers)
		assert.Len(t, excelData.Rows, 0)

		// Test adding a row
		err := excelData.AddRow([]interface{}{"John Doe", 30})
		assert.NoError(t, err)
		assert.Len(t, excelData.Rows, 1)
		assert.Equal(t, []interface{}{"John Doe", 30}, excelData.Rows[0])

		// Test writing to Excel
		err = excelData.ToExcel("test_new_excel_data.xlsx")
		assert.NoError(t, err)
		defer os.RemoveAll("test_new_excel_data.xlsx")

		// Verify file contents
		f, err := excelize.OpenFile("test_new_excel_data.xlsx")
		assert.NoError(t, err)
		defer f.Close()

		cells, err := f.GetRows("Sheet1")
		assert.NoError(t, err)
		assert.Len(t, cells, 2) // Header + 1 data row
		assert.Equal(t, headers, cells[0])
		assert.Equal(t, []string{"John Doe", "30"}, cells[1])
	})

	t.Run("NewExcelData with empty headers", func(t *testing.T) {
		excelData := NewExcelData[person]([]string{})
		assert.Len(t, excelData.Headers, 0)
	})

	t.Run("AddRow with mismatched column count", func(t *testing.T) {
		excelData := NewExcelData[person]([]string{"Name", "Age"})
		err := excelData.AddRow([]interface{}{"John"}) // Missing Age
		assert.Error(t, err)
		assert.Contains(t, err.Error(), "row length (1) does not match headers length (2)")
	})

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

func TestFromStructWithEmptySlice(t *testing.T) {
	var data []person

	_, err := FromStruct(data)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "input slice is empty")
}

func TestFromExcelWithNonExistentFile(t *testing.T) {
	_, err := FromExcel[person]("non_existent_file.xlsx")
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "no such file or directory")
}

func TestFromExcelWithEmptyFile(t *testing.T) {
	// Create an empty Excel file
	f := excelize.NewFile()
	inputPath := "./test_input/empty_file.xlsx"
	os.MkdirAll(filepath.Dir(inputPath), os.ModePerm)
	err := f.SaveAs(inputPath)
	assert.NoError(t, err)
	defer os.RemoveAll(filepath.Dir(inputPath))

	_, err = FromExcel[person](inputPath)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "excel file is empty or has no data rows")
}

func TestToStructWithTypeMismatch(t *testing.T) {
	type MismatchedPerson struct {
		Name string
		Age  string // Age is string instead of int
	}

	// Create test data
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Name")
	f.SetCellValue("Sheet1", "B1", "Age")
	f.SetCellValue("Sheet1", "A2", "John")
	f.SetCellValue("Sheet1", "B2", 30) // Age as int
	inputPath := "./mismatched_types.xlsx"
	err := f.SaveAs(inputPath)
	assert.NoError(t, err)
	defer os.Remove(inputPath)

	excelData, err := FromExcel[MismatchedPerson](inputPath)
	assert.NoError(t, err)

	result := excelData.ToStruct()
	assert.Len(t, result.Errors, 1)
	assert.Contains(t, result.Errors[0].Error(), "cannot convert '30' to type string")
	assert.Len(t, result.Data, 0) // No valid data due to type mismatch
}

func TestSetFieldWithUnsupportedType(t *testing.T) {
	type UnsupportedPerson struct {
		Name string
		Data []byte // Unsupported type
	}

	value := []byte("some data")
	field := reflect.ValueOf(&UnsupportedPerson{}).Elem().FieldByName("Data")

	err := setField(field, value)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "unsupported type: []uint8")
}

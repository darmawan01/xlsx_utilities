package xlsx_utilities

import (
	"reflect"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

func TestSetField(t *testing.T) {
	type TestStruct struct {
		String string
		Int    int
		Uint   uint
		Float  float64
		Bool   bool
		Time   time.Time
		Custom customID
	}

	tests := []struct {
		name      string
		fieldName string
		value     interface{}
		expected  interface{}
		expectErr bool
	}{
		{"String", "String", "test", "test", false},
		{"Int from int", "Int", 42, 42, false},
		{"Int from string", "Int", "42", 42, false},
		{"Uint from uint", "Uint", uint(42), uint(42), false},
		{"Uint from string", "Uint", "42", uint(42), false},
		{"Float from float", "Float", 3.14, 3.14, false},
		{"Float from string", "Float", "3.14", 3.14, false},
		{"Bool from bool", "Bool", true, true, false},
		{"Bool from string", "Bool", "true", true, false},
		{"Time from string", "Time", "2023-05-15T14:30:00Z", time.Date(2023, 5, 15, 14, 30, 0, 0, time.UTC), false},
		{"Custom type without parser", "Custom", "ID-00042", customID(0), true},
		{"Unsupported type", "String", []int{1, 2, 3}, nil, true},
	}

	delete(TypeConverters, reflect.TypeOf(customID(0)))
	delete(TypeParsers, reflect.TypeOf(customID(0)))

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			v := reflect.New(reflect.TypeOf(TestStruct{})).Elem()
			field := v.FieldByName(tt.fieldName)
			err := setField(field, tt.value)

			if tt.expectErr {
				assert.Error(t, err)
			} else {
				assert.NoError(t, err)
				assert.Equal(t, tt.expected, field.Interface())
			}
		})
	}
}

func TestToStructWithNestedAndArrays(t *testing.T) {
	type Address struct {
		Street string
		City   string
	}

	type Contact struct {
		Email string
		Phone string
	}

	type Person struct {
		Name     string
		Age      int
		Address  Address
		Contacts []Contact
	}

	headers := []string{
		"Name", "Age",
		"Address Street", "Address City",
		"Contacts Email", "Contacts Phone",
	}

	rows := [][]interface{}{
		{"Alice", 30, "123 Main St", "New York", "alice@example.com", "123-456-7890"},
		{"Bob", 25, "456 Elm St", "San Francisco", "bob@example.com", "987-654-3210"},
	}

	excelData := &ExcelData[*Person]{Headers: headers, Rows: rows}

	result := excelData.ToStruct()

	assert.Empty(t, result.Errors)
	assert.Len(t, result.Data, 2)

	// Check first person
	assert.Equal(t, "Alice", result.Data[0].Name)
	assert.Equal(t, 30, result.Data[0].Age)
	assert.Equal(t, "123 Main St", result.Data[0].Address.Street)
	assert.Equal(t, "New York", result.Data[0].Address.City)
	assert.Len(t, result.Data[0].Contacts, 1)
	assert.Equal(t, "alice@example.com", result.Data[0].Contacts[0].Email)
	assert.Equal(t, "123-456-7890", result.Data[0].Contacts[0].Phone)

	// Check second person
	assert.Equal(t, "Bob", result.Data[1].Name)
	assert.Equal(t, 25, result.Data[1].Age)
	assert.Equal(t, "456 Elm St", result.Data[1].Address.Street)
	assert.Equal(t, "San Francisco", result.Data[1].Address.City)
	assert.Len(t, result.Data[1].Contacts, 1)
	assert.Equal(t, "bob@example.com", result.Data[1].Contacts[0].Email)
	assert.Equal(t, "987-654-3210", result.Data[1].Contacts[0].Phone)
}

package xlsx_utilities

import (
	"fmt"
	"os"
	"reflect"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

type address struct {
	Street string
	City   string
}

type customID int

type user struct {
	Name      string
	Age       int
	Address   address
	BirthDate time.Time
	ID        customID
	id        int // Unexported field
}

func (c customID) String() string {
	return fmt.Sprintf("ID-%05d", int(c))
}

func TestXlsxUtilities(t *testing.T) {
	// Register custom type handlers
	RegisterTypeConverter(reflect.TypeOf(customID(0)), func(i interface{}) (string, error) {
		id, ok := i.(customID)
		if !ok {
			return "", fmt.Errorf("expected customID, got %T", i)
		}
		return id.String(), nil
	})

	RegisterTypeParser(reflect.TypeOf(customID(0)), func(s string) (interface{}, error) {
		var id int
		_, err := fmt.Sscanf(s, "ID-%05d", &id)
		if err != nil {
			return nil, err
		}
		return customID(id), nil
	})

	// Test data
	data := []user{
		{
			Name: "Alice",
			Age:  30,
			Address: address{
				Street: "123 Main St",
				City:   "New York",
			},
			BirthDate: time.Date(1990, 1, 1, 0, 0, 0, 0, time.UTC),
			ID:        customID(1),
			id:        1, // This will be ignored in Excel conversion
		},
		{
			Name: "Bob",
			Age:  25,
			Address: address{
				Street: "456 Elm St",
				City:   "San Francisco",
			},
			BirthDate: time.Date(1995, 5, 15, 0, 0, 0, 0, time.UTC),
			ID:        customID(2),
			id:        2, // This will be ignored in Excel conversion
		},
	}

	t.Run("FromStruct", func(t *testing.T) {
		excelData, err := FromStruct(data)
		assert.NoError(t, err)
		assert.Equal(t, []string{"Name", "Age", "Address Street", "Address City", "BirthDate", "ID"}, excelData.Headers)
		assert.Len(t, excelData.Rows, 2)
		assert.Equal(t, []interface{}{"Alice", 30, "123 Main St", "New York", "1990-01-01T00:00:00Z", "ID-00001"}, excelData.Rows[0])
	})

	t.Run("ToExcel and FromExcel", func(t *testing.T) {
		filename := "test_people.xlsx"
		defer os.Remove(filename) // Clean up after test

		// Convert to Excel
		excelData, err := FromStruct(data)
		assert.NoError(t, err)
		err = excelData.ToExcel(filename)
		assert.NoError(t, err)

		// Read from Excel
		readExcelData, err := FromExcel[user](filename)
		assert.NoError(t, err)
		assert.Equal(t, excelData.Headers, readExcelData.Headers)
		assert.Equal(t, excelData.Rows, readExcelData.Rows)
	})

	t.Run("ToStruct", func(t *testing.T) {
		excelData, err := FromStruct(data)
		assert.NoError(t, err)

		result := excelData.ToStruct()
		assert.Len(t, result.Data, 2)
		assert.Empty(t, result.Errors)

		// Check if the converted data matches the original data
		for i, person := range result.Data {
			assert.Equal(t, data[i].Name, person.Name)
			assert.Equal(t, data[i].Age, person.Age)
			assert.Equal(t, data[i].Address, person.Address)
			assert.True(t, data[i].BirthDate.Equal(person.BirthDate))
			assert.Equal(t, data[i].ID, person.ID)
			// Note: We don't check the 'id' field as it's unexported and should be ignored
		}
	})

	t.Run("Custom Type Handling", func(t *testing.T) {
		// Test customID handling
		id := customID(12345)
		converted, err := TypeConverters[reflect.TypeOf(id)](id)
		assert.NoError(t, err)
		assert.Equal(t, "ID-12345", converted)

		parsed, err := TypeParsers[reflect.TypeOf(id)]("ID-12345")
		assert.NoError(t, err)
		assert.Equal(t, customID(12345), parsed)
	})

	t.Run("Unexported Fields", func(t *testing.T) {
		excelData, err := FromStruct(data)
		assert.NoError(t, err)
		// Ensure that the 'id' field is not present in the headers
		assert.NotContains(t, excelData.Headers, "id")
	})
}

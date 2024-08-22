package xlsx_utilities

import (
	"os"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

type Address struct {
	Street *string
	City   string
	ZIP    *int
}

type ContactInfo struct {
	Email    *string
	Phone    string
	Address  *Address
	Verified *bool
}

type ComplexStruct struct {
	ID           *int
	Name         string
	BirthDate    *time.Time
	Height       *float64
	IsActive     *bool
	Tags         *[]string
	ContactInfos []*ContactInfo
}

func TestComplexPointerStruct(t *testing.T) {
	// Helper function to create string pointer
	strPtr := func(s string) *string { return &s }
	// Helper function to create int pointer
	intPtr := func(i int) *int { return &i }
	// Helper function to create float64 pointer
	float64Ptr := func(f float64) *float64 { return &f }
	// Helper function to create bool pointer
	boolPtr := func(b bool) *bool { return &b }
	// Helper function to create time.Time pointer
	timePtr := func(t time.Time) *time.Time { return &t }

	// Create test data
	testData := []*ComplexStruct{
		{
			ID:        intPtr(1),
			Name:      "Alice",
			BirthDate: timePtr(time.Date(1990, 1, 1, 0, 0, 0, 0, time.UTC)),
			Height:    float64Ptr(165.5),
			IsActive:  boolPtr(true),
			Tags:      &[]string{"tag1", "tag2"},
			ContactInfos: []*ContactInfo{
				{
					Email: strPtr("alice@example.com"),
					Phone: "123-456-7890",
					Address: &Address{
						Street: strPtr("123 Main St"),
						City:   "New York",
						ZIP:    intPtr(10001),
					},
					Verified: boolPtr(true),
				},
				{
					Email: strPtr("alice.work@example.com"),
					Phone: "098-765-4321",
					Address: &Address{
						Street: strPtr("456 Work Ave"),
						City:   "New York",
						ZIP:    intPtr(10002),
					},
					Verified: boolPtr(false),
				},
			},
		},
		{
			ID:        intPtr(2),
			Name:      "Bob",
			BirthDate: timePtr(time.Date(1985, 5, 15, 0, 0, 0, 0, time.UTC)),
			Height:    float64Ptr(180.0),
			IsActive:  boolPtr(false),
			Tags:      &[]string{"tag3", "tag4"},
			ContactInfos: []*ContactInfo{
				{
					Email: strPtr("bob@example.com"),
					Phone: "555-123-4567",
					Address: &Address{
						Street: strPtr("789 Oak Rd"),
						City:   "San Francisco",
						ZIP:    intPtr(94102),
					},
					Verified: boolPtr(true),
				},
			},
		},
	}

	t.Run("FromStruct with complex pointers", func(t *testing.T) {
		excelData, err := FromStruct(testData)
		assert.NoError(t, err)

		expectedHeaders := []string{
			"ID", "Name", "BirthDate", "Height", "IsActive",
			"Tags",
			"ContactInfos Email", "ContactInfos Phone", "ContactInfos Address Street", "ContactInfos Address City", "ContactInfos Address ZIP", "ContactInfos Verified",
		}
		assert.Equal(t, expectedHeaders, excelData.Headers)
		assert.Len(t, excelData.Rows, 2)

		// Check the first row
		assert.Len(t, excelData.Rows[0], len(expectedHeaders))
		assert.Equal(t, 1, excelData.Rows[0][0])
		assert.Equal(t, "Alice", excelData.Rows[0][1])
		assert.Equal(t, "1990-01-01T00:00:00Z", excelData.Rows[0][2])
		assert.Equal(t, 165.5, excelData.Rows[0][3])
		assert.Equal(t, true, excelData.Rows[0][4])

		excelData.Save("test.xlsx")
		// assert.Equal(t, "tag1", excelData.Rows[0][5])

		// Check first ContactInfo
		// assert.Equal(t, "alice@example.com", excelData.Rows[0][6])
		// assert.Equal(t, "123-456-7890", excelData.Rows[0][7])
		// assert.Equal(t, "123 Main St", excelData.Rows[0][8])
		// assert.Equal(t, "New York", excelData.Rows[0][9])
		// assert.Equal(t, 10001, excelData.Rows[0][10])
		// assert.Equal(t, true, excelData.Rows[0][11])
	})
	t.Run("Write to file and read back", func(t *testing.T) {
		filename := "test_complex_data.xlsx"
		defer os.Remove(filename) // Clean up after test

		// Write to file
		excelData, err := FromStruct(testData)
		assert.NoError(t, err)
		err = excelData.ToExcel(filename)
		assert.NoError(t, err)

		// Read from file
		readExcelData, err := FromExcel[*ComplexStruct](filename)
		assert.NoError(t, err)

		// Convert back to struct
		result := readExcelData.ToStruct()
		assert.Len(t, result.Data, len(testData))
		assert.Empty(t, result.Errors)

		// Verify data
		for i, original := range testData {
			if i < len(result.Data) {
				converted := result.Data[i]

				assert.Equal(t, *original.ID, *converted.ID)
				assert.Equal(t, original.Name, converted.Name)
				assert.True(t, original.BirthDate.Equal(*converted.BirthDate))
				assert.Equal(t, *original.Height, *converted.Height)
				assert.Equal(t, *original.IsActive, *converted.IsActive)
				assert.Equal(t, *original.Tags, *converted.Tags)

				assert.Len(t, converted.ContactInfos, len(original.ContactInfos))
				for j, originalContact := range original.ContactInfos {
					if j < len(converted.ContactInfos) {
						convertedContact := converted.ContactInfos[j]

						assert.Equal(t, *originalContact.Email, *convertedContact.Email)
						assert.Equal(t, originalContact.Phone, convertedContact.Phone)
						assert.Equal(t, *originalContact.Address.Street, *convertedContact.Address.Street)
						assert.Equal(t, originalContact.Address.City, convertedContact.Address.City)
						assert.Equal(t, *originalContact.Address.ZIP, *convertedContact.Address.ZIP)
						assert.Equal(t, *originalContact.Verified, *convertedContact.Verified)
					}
				}
			}
		}
	})
}

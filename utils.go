package xlsx_utilities

import "strconv"

// convertCellValue attempts to convert string cell values to appropriate types
func convertCellValue(value string) interface{} {
	// Try converting to int
	if intVal, err := strconv.Atoi(value); err == nil {
		return intVal
	}

	// Try converting to float
	if floatVal, err := strconv.ParseFloat(value, 64); err == nil {
		return floatVal
	}

	// Try converting to bool
	if boolVal, err := strconv.ParseBool(value); err == nil {
		return boolVal
	}

	// If all else fails, return as string
	return value
}

package xlsx_utilities

import "strconv"

// convertCellValue attempts to convert string cell values to appropriate types
func convertCellValue(value string) interface{} {
	if intVal, err := strconv.Atoi(value); err == nil {
		return intVal
	}
	if floatVal, err := strconv.ParseFloat(value, 64); err == nil {
		return floatVal
	}
	if boolVal, err := strconv.ParseBool(value); err == nil {
		return boolVal
	}
	return value // If not a number or bool, return as string
}

package xlsx_utilities

import (
	"fmt"
	"reflect"
	"time"
)

// CustomTypeConverter is a function type for custom type conversions
type CustomTypeConverter func(interface{}) (string, error)

// CustomTypeParser is a function type for parsing custom types from strings
type CustomTypeParser func(string) (interface{}, error)

// TypeConverters maps types to their custom converters
var TypeConverters = map[reflect.Type]CustomTypeConverter{}

// TypeParsers maps types to their custom parsers
var TypeParsers = map[reflect.Type]CustomTypeParser{}

// RegisterTypeConverter registers a custom type converter
func RegisterTypeConverter(t reflect.Type, converter CustomTypeConverter) {
	TypeConverters[t] = converter
}

// RegisterTypeParser registers a custom type parser
func RegisterTypeParser(t reflect.Type, parser CustomTypeParser) {
	TypeParsers[t] = parser
}

// init function to register built-in custom type handlers
func init() {
	// Register time.Time handlers
	RegisterTypeConverter(reflect.TypeOf(time.Time{}), func(i interface{}) (string, error) {
		t, ok := i.(time.Time)
		if !ok {
			return "", fmt.Errorf("expected time.Time, got %T", i)
		}
		return t.Format(time.RFC3339), nil
	})

	RegisterTypeParser(reflect.TypeOf(time.Time{}), func(s string) (interface{}, error) {
		return time.Parse(time.RFC3339, s)
	})
}

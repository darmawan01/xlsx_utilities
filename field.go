package xlsx_utilities

import (
	"fmt"
	"reflect"
	"strconv"
	"time"
)

// setField sets the value of a struct field, handling type conversions
func setField(field reflect.Value, value interface{}) error {
	// Handle pointer types
	if field.Kind() == reflect.Ptr {
		if field.IsNil() {
			field.Set(reflect.New(field.Type().Elem()))
		}
		field = field.Elem()
	}

	// Check for custom type parser
	if parser, ok := TypeParsers[field.Type()]; ok {
		parsedValue, err := parser(fmt.Sprintf("%v", value))
		if err != nil {
			return fmt.Errorf("error parsing custom type: %v", err)
		}
		field.Set(reflect.ValueOf(parsedValue))
		return nil
	}

	// Handle built-in types
	switch field.Kind() {
	case reflect.String:
		s, err := toString(value)
		if err != nil {
			return err
		}
		field.SetString(s)
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		i, err := toInt64(value)
		if err != nil {
			return err
		}
		field.SetInt(i)
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		u, err := toUint64(value)
		if err != nil {
			return err
		}
		field.SetUint(u)
	case reflect.Float32, reflect.Float64:
		f, err := toFloat64(value)
		if err != nil {
			return err
		}
		field.SetFloat(f)
	case reflect.Bool:
		b, err := toBool(value)
		if err != nil {
			return err
		}
		field.SetBool(b)
	case reflect.Struct:
		// Handle time.Time as a special case
		if field.Type() == reflect.TypeOf(time.Time{}) {
			t, err := toTime(value)
			if err != nil {
				return err
			}
			field.Set(reflect.ValueOf(t))
		} else {
			return fmt.Errorf("unsupported struct type: %v", field.Type())
		}
	default:
		return fmt.Errorf("unsupported type: %v", field.Type())
	}
	return nil
}

// Add this new helper function
func toString(value interface{}) (string, error) {
	switch v := value.(type) {
	case string:
		return v, nil
	case fmt.Stringer:
		return v.String(), nil
	default:
		return fmt.Sprintf("%v", v), nil
	}
}

// Helper functions for type conversion
func toInt64(value interface{}) (int64, error) {
	switch v := value.(type) {
	case int:
		return int64(v), nil
	case int64:
		return v, nil
	case float64:
		return int64(v), nil
	case string:
		return strconv.ParseInt(v, 10, 64)
	default:
		return 0, fmt.Errorf("cannot convert %v to int64", value)
	}
}

func toUint64(value interface{}) (uint64, error) {
	switch v := value.(type) {
	case uint:
		return uint64(v), nil
	case uint64:
		return v, nil
	case float64:
		if v < 0 {
			return 0, fmt.Errorf("cannot convert negative float64 to uint64")
		}
		return uint64(v), nil
	case string:
		return strconv.ParseUint(v, 10, 64)
	default:
		return 0, fmt.Errorf("cannot convert %v to uint64", value)
	}
}

func toFloat64(value interface{}) (float64, error) {
	switch v := value.(type) {
	case float64:
		return v, nil
	case int:
		return float64(v), nil
	case int64:
		return float64(v), nil
	case string:
		return strconv.ParseFloat(v, 64)
	default:
		return 0, fmt.Errorf("cannot convert %v to float64", value)
	}
}

func toBool(value interface{}) (bool, error) {
	switch v := value.(type) {
	case bool:
		return v, nil
	case string:
		return strconv.ParseBool(v)
	case int:
		return v != 0, nil
	default:
		return false, fmt.Errorf("cannot convert %v to bool", value)
	}
}

func toTime(value interface{}) (time.Time, error) {
	switch v := value.(type) {
	case time.Time:
		return v, nil
	case string:
		return time.Parse(time.RFC3339, v)
	default:
		return time.Time{}, fmt.Errorf("cannot convert %v to time.Time", value)
	}
}

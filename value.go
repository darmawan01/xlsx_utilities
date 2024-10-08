package xlsx_utilities

import (
	"fmt"
	"reflect"
	"time"
)

func getStructValues(v reflect.Value) ([]interface{}, error) {
	return getNestedValues(v)
}

func getNestedValues(v reflect.Value) ([]interface{}, error) {
	if v.Kind() == reflect.Ptr {
		if v.IsNil() {
			return []interface{}{getDefaultValue(v.Type().Elem())}, nil
		}
		v = v.Elem()
	}

	if converter, ok := TypeConverters[v.Type()]; ok {
		converted, err := converter(v.Interface())
		if err != nil {
			return nil, fmt.Errorf("error converting custom type: %v", err)
		}
		return []interface{}{converted}, nil
	}

	if v.Kind() != reflect.Struct {
		if v.Kind() == reflect.Slice {
			return getFirstSliceValue(v)
		}
		return []interface{}{v.Interface()}, nil
	}

	var values []interface{}

	for i := 0; i < v.NumField(); i++ {
		field := v.Field(i)
		fieldType := v.Type().Field(i)

		if !fieldType.IsExported() {
			continue
		}

		if converter, ok := TypeConverters[field.Type()]; ok {
			converted, err := converter(field.Interface())
			if err != nil {
				return nil, fmt.Errorf("error converting custom type: %v", err)
			}
			values = append(values, converted)
			continue
		}

		switch field.Kind() {
		case reflect.Ptr:
			if field.IsNil() {
				values = append(values, getDefaultValue(field.Type().Elem()))
			} else {
				nestedValues, err := getNestedValues(field.Elem())
				if err != nil {
					return nil, err
				}
				values = append(values, nestedValues...)
			}
		case reflect.Struct:
			if field.Type() == reflect.TypeOf(time.Time{}) {
				values = append(values, field.Interface())
			} else {
				nestedValues, err := getNestedValues(field)
				if err != nil {
					return nil, err
				}
				values = append(values, nestedValues...)
			}
		case reflect.Slice:
			sliceValues, err := getFirstSliceValue(field)
			if err != nil {
				return nil, err
			}
			values = append(values, sliceValues...)
		default:
			values = append(values, field.Interface())
		}
	}

	return values, nil
}

func getFirstSliceValue(slice reflect.Value) ([]interface{}, error) {
	if slice.Len() == 0 {
		return []interface{}{getDefaultValue(slice.Type().Elem())}, nil
	}

	firstElem := slice.Index(0)
	if firstElem.Kind() == reflect.Ptr {
		if firstElem.IsNil() {
			return []interface{}{getDefaultValue(firstElem.Type().Elem())}, nil
		}
		firstElem = firstElem.Elem()
	}

	if firstElem.Kind() == reflect.Struct {
		// Flatten the first struct element
		flattenedValues, err := getNestedValues(firstElem)
		if err != nil {
			return nil, err
		}

		emptyValues := make([]interface{}, 0, len(flattenedValues))
		// Pad the rest of the slice with empty strings
		for i := 0; i < len(flattenedValues); i++ {
			emptyValues = append(emptyValues, "")
		}
		return emptyValues, nil
	}

	return []interface{}{""}, nil
}

func getDefaultValue(t reflect.Type) interface{} {
	switch t.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return int64(0)
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		return uint64(0)
	case reflect.Float32, reflect.Float64:
		return float64(0)
	case reflect.Bool:
		return false
	case reflect.String:
		return ""
	default:
		return nil
	}
}

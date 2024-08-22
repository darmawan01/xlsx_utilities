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
			return []interface{}{nil}, nil
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
				values = append(values, nil)
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
		return []interface{}{nil}, nil
	}

	firstElem := slice.Index(0)

	// Handle pointer to struct
	if firstElem.Kind() == reflect.Ptr {
		if firstElem.IsNil() {
			return []interface{}{nil}, nil
		}
		firstElem = firstElem.Elem()
	}

	// If it's a struct, we need to flatten it
	if firstElem.Kind() == reflect.Struct {
		return getNestedValues(firstElem)
	}

	// For non-struct types, just return the value
	return []interface{}{firstElem.Interface()}, nil
}

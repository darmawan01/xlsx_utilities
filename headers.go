package xlsx_utilities

import (
	"reflect"
	"time"
)

func getStructHeaders(t reflect.Type) ([]string, error) {
	return getNestedHeaders(t, "")
}

func getNestedHeaders(t reflect.Type, prefix string) ([]string, error) {
	if t.Kind() == reflect.Ptr {
		t = t.Elem()
	}

	if _, ok := TypeConverters[t]; ok {
		return []string{prefix}, nil
	}

	if t.Kind() != reflect.Struct {
		return []string{prefix}, nil
	}

	var headers []string

	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)

		if !field.IsExported() {
			continue
		}

		fieldName := field.Name
		if prefix != "" {
			fieldName = prefix + " " + fieldName
		}

		fieldType := field.Type
		if fieldType.Kind() == reflect.Ptr {
			fieldType = fieldType.Elem()
		}

		switch fieldType.Kind() {
		case reflect.Struct:
			if fieldType == reflect.TypeOf(time.Time{}) {
				headers = append(headers, fieldName)
			} else {
				nestedHeaders, err := getNestedHeaders(fieldType, fieldName)
				if err != nil {
					return nil, err
				}
				headers = append(headers, nestedHeaders...)
			}
		case reflect.Slice:
			sliceElemType := fieldType.Elem()
			if sliceElemType.Kind() == reflect.Ptr {
				sliceElemType = sliceElemType.Elem()
			}
			nestedHeaders, err := getNestedHeaders(sliceElemType, fieldName)
			if err != nil {
				return nil, err
			}
			headers = append(headers, nestedHeaders...)
		default:
			headers = append(headers, fieldName)
		}
	}

	return headers, nil
}

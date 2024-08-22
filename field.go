package xlsx_utilities

import (
	"errors"
	"fmt"
	"reflect"
	"strconv"
	"strings"
	"time"
)

func setNestedField(v reflect.Value, fieldPath string, value interface{}) error {
	fields := strings.Split(fieldPath, " ")
	for i, field := range fields {
		if v.Kind() == reflect.Ptr {
			if v.IsNil() {
				v.Set(reflect.New(v.Type().Elem()))
			}
			v = v.Elem()
		}

		if v.Kind() != reflect.Struct {
			return fmt.Errorf("not a struct: %v", v.Kind())
		}

		f := v.FieldByName(field)
		if !f.IsValid() {
			return fmt.Errorf("no such field: %s in obj", field)
		}

		if i == len(fields)-1 {
			// Check if there's a custom type converter
			if converter, ok := TypeParsers[f.Type()]; ok {
				convertedValue, err := converter(fmt.Sprintf("%v", value))
				if err != nil {
					return fmt.Errorf("error parsing custom type: %v", err)
				}
				f.Set(reflect.ValueOf(convertedValue))
				return nil
			}

			return setField(f, value)
		}

		if f.Kind() == reflect.Ptr {
			if f.IsNil() {
				f.Set(reflect.New(f.Type().Elem()))
			}
			v = f.Elem()
		} else if f.Kind() == reflect.Slice {
			if f.IsNil() {
				f.Set(reflect.MakeSlice(f.Type(), 0, 0))
			}
			newElem := reflect.New(f.Type().Elem()).Elem()
			f.Set(reflect.Append(f, newElem))
			v = f.Index(f.Len() - 1)
		} else {
			v = f
		}
	}
	return nil
}

// setField sets the value of a struct field, handling type conversions
func setField(field reflect.Value, value interface{}) error {
	if !field.CanSet() {
		return fmt.Errorf("cannot set field")
	}

	// Handle pointer types
	if field.Kind() == reflect.Ptr {
		if field.IsNil() {
			field.Set(reflect.New(field.Type().Elem()))
		}
		return setField(field.Elem(), value)
	}

	// Check if there's a custom type converter
	if converter, ok := TypeParsers[field.Type()]; ok {
		convertedValue, err := converter(fmt.Sprintf("%v", value))
		if err != nil {
			return fmt.Errorf("error parsing custom type: %v", err)
		}
		field.Set(reflect.ValueOf(convertedValue))
		return nil
	}

	val := reflect.ValueOf(value)

	switch field.Kind() {
	case reflect.String:
		if val.Kind() == reflect.String {
			field.SetString(val.String())
		} else {
			return errors.New("value is not a string")
		}
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		intVal, err := strconv.ParseInt(fmt.Sprintf("%v", value), 10, 64)
		if err != nil {
			return err
		}
		field.SetInt(intVal)
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		uintVal, err := strconv.ParseUint(fmt.Sprintf("%v", value), 10, 64)
		if err != nil {
			return err
		}
		field.SetUint(uintVal)
	case reflect.Float32, reflect.Float64:
		floatVal, err := strconv.ParseFloat(fmt.Sprintf("%v", value), 64)
		if err != nil {
			return err
		}
		field.SetFloat(floatVal)
	case reflect.Bool:
		boolVal, err := strconv.ParseBool(fmt.Sprintf("%v", value))
		if err != nil {
			return err
		}
		field.SetBool(boolVal)
	case reflect.Struct:
		if field.Type() == reflect.TypeOf(time.Time{}) {
			timeVal, err := time.Parse(time.RFC3339, fmt.Sprintf("%v", value))
			if err != nil {
				return err
			}
			field.Set(reflect.ValueOf(timeVal))
		} else {
			// For other struct types, we'll assume the value is a map
			if mapValue, ok := value.(map[string]interface{}); ok {
				for key, val := range mapValue {
					err := setNestedField(field, key, val)
					if err != nil {
						return err
					}
				}
			} else {
				return fmt.Errorf("unsupported struct type: %v", field.Type())
			}
		}
	case reflect.Slice:
		return setSliceField(field, value)
	default:
		return fmt.Errorf("unsupported type: %v", field.Type())
	}

	return nil
}

func setSliceField(field reflect.Value, value interface{}) error {
	sliceValue := reflect.ValueOf(value)
	if sliceValue.Kind() != reflect.Slice {
		return fmt.Errorf("expected slice, got %v", sliceValue.Kind())
	}

	sliceType := field.Type()
	slice := reflect.MakeSlice(sliceType, sliceValue.Len(), sliceValue.Len())

	for i := 0; i < sliceValue.Len(); i++ {
		err := setField(slice.Index(i), sliceValue.Index(i).Interface())
		if err != nil {
			return err
		}
	}

	field.Set(slice)
	return nil
}

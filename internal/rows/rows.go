// Package rows maps struct fields to spreadsheet columns. It is shared by the
// query package's structured reader (ReadRows) and the editor package's
// structured writer (WriteRows) so both sides agree on the exact same mapping:
// the `excel:"..."` struct tag selects the column name, a `excel:"-"` tag
// skips the field, and untagged fields map to their own name.
package rows

import (
	"fmt"
	"reflect"
	"strings"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
)

// Column is one struct field mapped to a spreadsheet column.
type Column struct {
	// Name is the resolved column name: the `excel` tag value when present,
	// otherwise the struct field name.
	Name string
	// Index is the position of the field in the struct's top-level fields.
	Index int
	// Type is the field's type.
	Type reflect.Type
}

// Columns validates t as a table row type and returns its ordered columns.
//
// t must be a struct or a pointer to a struct. Exported top-level fields are
// mapped by their `excel:"name"` tag, or by the field name when no tag is
// present; unexported fields and fields tagged `excel:"-"` are skipped.
// Embedded fields are not flattened. A type with no mapped fields is rejected.
func Columns(t reflect.Type) ([]Column, error) {
	if t.Kind() == reflect.Ptr {
		t = t.Elem()
	}
	if t.Kind() != reflect.Struct {
		return nil, fmt.Errorf("row type must be a struct, got %s: %w", t, toolkiterrors.ErrInvalidValue)
	}
	cols := make([]Column, 0, t.NumField())
	for i := 0; i < t.NumField(); i++ {
		f := t.Field(i)
		if !f.IsExported() {
			continue
		}
		name := columnName(f)
		if name == "" {
			continue // `excel:"-"`
		}
		cols = append(cols, Column{Name: name, Index: i, Type: f.Type})
	}
	if len(cols) == 0 {
		return nil, fmt.Errorf("row type %s has no excel-mapped fields: %w", t, toolkiterrors.ErrInvalidValue)
	}
	return cols, nil
}

// columnName resolves a field's column name from its excel tag. A missing or
// empty tag falls back to the field name; the tag "-" means "skip".
func columnName(f reflect.StructField) string {
	tag, ok := f.Tag.Lookup("excel")
	if !ok {
		return f.Name
	}
	tag = strings.TrimSpace(tag)
	switch tag {
	case "-":
		return ""
	case "":
		return f.Name
	default:
		return tag
	}
}

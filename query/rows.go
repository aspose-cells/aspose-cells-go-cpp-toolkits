package query

import (
	"fmt"
	"reflect"
	"strings"
	"time"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/rows"
)

var timeType = reflect.TypeOf(time.Time{})

// ReadRows reads a worksheet's used range into a slice of T, mapping each
// struct field to a column via the header row.
//
// The first (header) row names the columns; each struct field is matched
// against a header cell — case-insensitively, after trimming whitespace — by
// its `excel:"name"` tag, or by its own name when no tag is present. Every
// mapped field must find a matching header column: an unknown or missing
// column returns ErrColumnNotFound (add `excel:"-"` to ignore a field, or a
// tag to rename it). Columns present in the header but not in the struct are
// ignored. Data rows below the header are decoded into T; an empty cell leaves
// the field at its zero value.
//
// T must be a struct. Supported field types are string, all integer and
// unsigned integer sizes, float32/float64, bool, and time.Time. A value whose
// kind does not match the field type, or a number that overflows it, returns an
// error wrapped in ErrInvalidValue.
//
// Example:
//
//	type Employee struct {
//		ID   int       `excel:"id"`
//		Name string    `excel:"name"`
//		Hire time.Time `excel:"hired_on"`
//	}
//	rows, err := query.ReadRows[Employee](
//		datasource.FilePathSource("employees.xlsx"),
//		query.WithSheetIndex(0),
//	)
func ReadRows[T any](source datasource.DataSource, opts ...Option) ([]T, error) {
	var zero T
	typ := reflect.TypeOf(zero)
	if typ.Kind() != reflect.Struct {
		return nil, fmt.Errorf("ReadRows: T must be a struct, got %s: %w", typ, toolkiterrors.ErrInvalidValue)
	}
	cols, err := rows.Columns(typ)
	if err != nil {
		return nil, err
	}
	grid, err := ReadWorksheet(source, opts...)
	if err != nil {
		return nil, err
	}
	if len(grid) == 0 {
		return []T{}, nil // empty sheet: no header, no rows
	}
	colPos, err := headerPositions(grid[0], cols)
	if err != nil {
		return nil, err
	}
	out := make([]T, 0, len(grid)-1)
	for r := 1; r < len(grid); r++ {
		var item T
		rv := reflect.ValueOf(&item).Elem()
		for i, c := range cols {
			cell := grid[r][colPos[i]]
			if err := assignCell(rv.Field(c.Index), cell); err != nil {
				return nil, fmt.Errorf("row %d, column %q: %w", r+1, c.Name, err)
			}
		}
		out = append(out, item)
	}
	return out, nil
}

// headerPositions locates each mapped column in the header row, returning the
// grid column index per mapped field in the same order as cols.
func headerPositions(header []CellValue, cols []rows.Column) ([]int, error) {
	pos := make([]int, len(cols))
	for i, c := range cols {
		found := -1
		for col, cell := range header {
			s, ok := cell.String()
			if ok && strings.EqualFold(strings.TrimSpace(s), c.Name) {
				found = col
				break
			}
		}
		if found < 0 {
			return nil, fmt.Errorf("column %q not found in header row: %w", c.Name, toolkiterrors.ErrColumnNotFound)
		}
		pos[i] = found
	}
	return pos, nil
}

// assignCell decodes a cell value into the reflect field fv. Empty cells leave
// the field untouched (zero value). A value of an incompatible kind or one that
// overflows the field type is an error.
func assignCell(fv reflect.Value, cv CellValue) error {
	if cv.IsEmpty() {
		return nil
	}
	switch fv.Kind() {
	case reflect.String:
		s, ok := cv.String()
		if !ok {
			return cellMismatch(cv, fv.Type())
		}
		fv.SetString(s)
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		n, ok := cv.Int()
		if !ok {
			return cellMismatch(cv, fv.Type())
		}
		if fv.OverflowInt(n) {
			return cellOverflow(cv, fv.Type(), n)
		}
		fv.SetInt(n)
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		n, ok := cv.Int()
		if !ok || n < 0 {
			return cellMismatch(cv, fv.Type())
		}
		if fv.OverflowUint(uint64(n)) {
			return cellOverflow(cv, fv.Type(), n)
		}
		fv.SetUint(uint64(n))
	case reflect.Float32, reflect.Float64:
		f, ok := cv.Float()
		if !ok {
			// A whole-number cell is KindInt; accept it into a float field too.
			if n, ok2 := cv.Int(); ok2 {
				f, ok = float64(n), true
			}
			if !ok {
				return cellMismatch(cv, fv.Type())
			}
		}
		if fv.OverflowFloat(f) {
			return cellOverflow(cv, fv.Type(), f)
		}
		fv.SetFloat(f)
	case reflect.Bool:
		b, ok := cv.Bool()
		if !ok {
			return cellMismatch(cv, fv.Type())
		}
		fv.SetBool(b)
	default:
		if fv.Type() == timeType {
			t, ok := cv.Time()
			if !ok {
				return cellMismatch(cv, fv.Type())
			}
			fv.Set(reflect.ValueOf(t))
			return nil
		}
		return fmt.Errorf("unsupported field type %s: %w", fv.Type(), toolkiterrors.ErrInvalidValue)
	}
	return nil
}

func cellMismatch(cv CellValue, t reflect.Type) error {
	return fmt.Errorf("cell of kind %s cannot be read as %s: %w", cv.Kind(), t, toolkiterrors.ErrInvalidValue)
}

func cellOverflow(cv CellValue, t reflect.Type, v any) error {
	return fmt.Errorf("cell value %v does not fit in %s: %w", v, t, toolkiterrors.ErrInvalidValue)
}

package query

import (
	"math"
	"strconv"
	"time"

	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// CellKind distinguishes the semantic types of a cell value. Callers switch on
// Kind() and then read the value with the matching accessor, instead of
// relying on reflection or string parsing.
type CellKind uint8

const (
	// KindEmpty is an empty cell, including the non-anchor cells of a merged
	// region.
	KindEmpty CellKind = iota
	// KindText is a string cell.
	KindText
	// KindInt is a numeric cell whose value is a whole number.
	KindInt
	// KindFloat is a numeric cell whose value is not a whole number.
	KindFloat
	// KindBool is a boolean cell.
	KindBool
	// KindDateTime is a date/time cell.
	KindDateTime
	// KindError is a cell holding a formula error (e.g. "#DIV/0!").
	KindError
)

// CellValue is a typed cell value. It pairs a value with its CellKind, so a
// program can inspect Kind() and read the value with the matching accessor.
// Accessors return (zero, false) when the value's kind does not match the
// requested type.
type CellValue struct {
	kind CellKind
	s    string
	i    int64
	f    float64
	b    bool
	t    time.Time
}

// Kind returns the semantic type of the value.
func (v CellValue) Kind() CellKind { return v.kind }

// IsEmpty reports whether the cell holds no value.
func (v CellValue) IsEmpty() bool { return v.kind == KindEmpty }

// String returns the value in its canonical text form for any non-empty kind,
// or ("", false) for an empty cell. For KindText and KindError it is the raw
// cell text; for numbers, booleans, and dates it is a formatted rendering.
func (v CellValue) String() (string, bool) {
	switch v.kind {
	case KindText, KindError:
		return v.s, true
	case KindInt:
		return strconv.FormatInt(v.i, 10), true
	case KindFloat:
		return strconv.FormatFloat(v.f, 'g', -1, 64), true
	case KindBool:
		return strconv.FormatBool(v.b), true
	case KindDateTime:
		return v.t.Format(time.RFC3339), true
	default:
		return "", false
	}
}

// Int returns the value as int64 for KindInt cells.
func (v CellValue) Int() (int64, bool) {
	if v.kind == KindInt {
		return v.i, true
	}
	return 0, false
}

// Float returns the value as float64 for KindFloat cells.
func (v CellValue) Float() (float64, bool) {
	if v.kind == KindFloat {
		return v.f, true
	}
	return 0, false
}

// Bool returns the value for KindBool cells.
func (v CellValue) Bool() (bool, bool) {
	if v.kind == KindBool {
		return v.b, true
	}
	return false, false
}

// Time returns the value as time.Time for KindDateTime cells.
func (v CellValue) Time() (time.Time, bool) {
	if v.kind == KindDateTime {
		return v.t, true
	}
	return time.Time{}, false
}

// fromCell maps an engine cell to a CellValue based on its CellValueType.
func fromCell(cell *asposecells.Cell) (CellValue, error) {
	cellType, err := cell.GetType()
	if err != nil {
		return CellValue{}, err
	}
	switch cellType {
	case asposecells.CellValueType_IsNull:
		return CellValue{kind: KindEmpty}, nil
	case asposecells.CellValueType_IsString:
		s, err := cell.GetStringValue()
		if err != nil {
			return CellValue{}, err
		}
		return CellValue{kind: KindText, s: s}, nil
	case asposecells.CellValueType_IsBool:
		b, err := cell.GetBoolValue()
		if err != nil {
			return CellValue{}, err
		}
		return CellValue{kind: KindBool, b: b}, nil
	case asposecells.CellValueType_IsDateTime:
		t, err := cell.GetDateTimeValue()
		if err != nil {
			return CellValue{}, err
		}
		return CellValue{kind: KindDateTime, t: t}, nil
	case asposecells.CellValueType_IsNumeric:
		// Dates are stored as serial numbers, so a numeric cell whose number
		// format is a date/time format is a date. The engine reports it via the
		// cell's style rather than the value type.
		style, err := cell.GetStyle()
		if err != nil {
			return CellValue{}, err
		}
		isDateTime, err := style.IsDateTime()
		if err != nil {
			return CellValue{}, err
		}
		if isDateTime {
			t, err := cell.GetDateTimeValue()
			if err != nil {
				return CellValue{}, err
			}
			return CellValue{kind: KindDateTime, t: t}, nil
		}
		v, err := cell.GetDoubleValue()
		if err != nil {
			return CellValue{}, err
		}
		return numericValue(v), nil
	case asposecells.CellValueType_IsError:
		s, err := cell.GetStringValue()
		if err != nil {
			return CellValue{}, err
		}
		return CellValue{kind: KindError, s: s}, nil
	default: // CellValueType_IsUnknown
		s, err := cell.GetStringValue()
		if err != nil {
			return CellValue{}, err
		}
		if s == "" {
			return CellValue{kind: KindEmpty}, nil
		}
		return CellValue{kind: KindText, s: s}, nil
	}
}

// numericValue classifies a numeric engine value as an integer when it is a
// whole number within int64 range, otherwise as a float.
func numericValue(v float64) CellValue {
	if v == math.Trunc(v) && v >= math.MinInt64 && v <= math.MaxInt64 {
		return CellValue{kind: KindInt, i: int64(v)}
	}
	return CellValue{kind: KindFloat, f: v}
}

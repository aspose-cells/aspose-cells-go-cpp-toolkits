package cells

import (
	"fmt"
	"strconv"
	"strings"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
)

// CellRef is a zero-based cell coordinate. String renders it in Excel style,
// so (Row: 2, Col: 1) renders as "B3". It is a plain value type, not an
// engine object, and is re-exported by the query package as its public
// addressing type.
type CellRef struct {
	Row int
	Col int
}

// String renders the reference in Excel style, e.g. (Row: 2, Col: 1) -> "B3".
func (r CellRef) String() string {
	return columnName(r.Col+1) + strconv.Itoa(r.Row+1)
}

// AbsoluteString renders the reference with absolute row and column markers,
// e.g. (Row: 2, Col: 1) -> "$B$3". This is the form named-range references
// use in the refersTo text the engine writes.
func (r CellRef) AbsoluteString() string {
	return "$" + columnName(r.Col+1) + "$" + strconv.Itoa(r.Row+1)
}

// Area is a rectangular cell region with inclusive start and end cells.
type Area struct {
	Start CellRef
	End   CellRef
}

// String renders the area in Excel style, e.g. "A1:C3".
func (a Area) String() string {
	return a.Start.String() + ":" + a.End.String()
}

// columnName converts a one-based column number to its Excel letters, e.g.
// 1 -> "A", 26 -> "Z", 27 -> "AA".
func columnName(n int) string {
	var b []byte
	for n > 0 {
		n--
		b = append(b, byte('A'+n%26))
		n /= 26
	}
	for i, j := 0, len(b)-1; i < j; i, j = i+1, j-1 {
		b[i], b[j] = b[j], b[i]
	}
	return string(b)
}

// columnIndex converts Excel column letters to a zero-based index, e.g.
// "A" -> 0, "Z" -> 25, "AA" -> 26. Non-letter input is an error.
func columnIndex(letters string) (int, error) {
	col := 0
	for _, ch := range letters {
		switch {
		case ch >= 'a' && ch <= 'z':
			ch -= 'a' - 'A'
		case ch >= 'A' && ch <= 'Z':
		default:
			return 0, fmt.Errorf("invalid column %q: %w", letters, toolkiterrors.ErrInvalidCellRef)
		}
		col = col*26 + int(ch-'A') + 1
	}
	return col - 1, nil
}

func isAlpha(b byte) bool {
	return (b >= 'a' && b <= 'z') || (b >= 'A' && b <= 'Z')
}

// ParseCellRef parses an Excel-style cell reference such as "B3" into a
// zero-based CellRef. Column letters are case-insensitive. Empty input, a
// reference without letters or without a row number, and non-numeric rows all
// return ErrInvalidCellRef.
func ParseCellRef(s string) (CellRef, error) {
	i := 0
	for i < len(s) && isAlpha(s[i]) {
		i++
	}
	if i == 0 || i == len(s) {
		return CellRef{}, fmt.Errorf("invalid cell reference %q: %w", s, toolkiterrors.ErrInvalidCellRef)
	}
	col, err := columnIndex(s[:i])
	if err != nil {
		return CellRef{}, fmt.Errorf("invalid cell reference %q: %w", s, toolkiterrors.ErrInvalidCellRef)
	}
	row1, err := strconv.Atoi(s[i:])
	if err != nil || row1 < 1 {
		return CellRef{}, fmt.Errorf("invalid cell reference %q: %w", s, toolkiterrors.ErrInvalidCellRef)
	}
	return CellRef{Row: row1 - 1, Col: col}, nil
}

// ParseArea parses an Excel-style area such as "A1:C3" into an Area. A single
// cell ("B2") is accepted as a one-cell area. Areas whose start cell lies below
// or to the right of the end cell, or with more than one ":" separator, return
// ErrInvalidRange.
func ParseArea(s string) (Area, error) {
	parts := strings.Split(s, ":")
	if len(parts) > 2 {
		return Area{}, fmt.Errorf("invalid area %q: %w", s, toolkiterrors.ErrInvalidRange)
	}
	start, err := ParseCellRef(parts[0])
	if err != nil {
		return Area{}, err
	}
	end := start
	if len(parts) == 2 {
		end, err = ParseCellRef(parts[1])
		if err != nil {
			return Area{}, err
		}
	}
	if start.Row > end.Row || start.Col > end.Col {
		return Area{}, fmt.Errorf("invalid area %q: %w", s, toolkiterrors.ErrInvalidRange)
	}
	return Area{Start: start, End: end}, nil
}

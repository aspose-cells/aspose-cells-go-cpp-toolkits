package query

import (
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
)

// CellRef is a zero-based cell coordinate. String renders it in Excel style
// (e.g. (Row: 2, Col: 1) -> "B3"), and ParseCellRef parses Excel references
// back into a CellRef.
type CellRef = cells.CellRef

// Area is a rectangular cell region with inclusive start and end cells.
type Area = cells.Area

// ParseCellRef parses an Excel-style cell reference such as "B3" into a
// zero-based CellRef. Column letters are case-insensitive.
func ParseCellRef(s string) (CellRef, error) { return cells.ParseCellRef(s) }

// ParseArea parses an Excel-style area such as "A1:C3" into an Area. A single
// cell ("B2") is accepted as a one-cell area.
func ParseArea(s string) (Area, error) { return cells.ParseArea(s) }

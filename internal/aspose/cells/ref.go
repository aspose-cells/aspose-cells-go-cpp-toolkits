package cells

import (
	refs "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/refs"
)

// The addressing vocabulary lives in internal/aspose/refs, a leaf package the
// save options can reach as well; a save option cannot import this package,
// because this package reaches formats, and formats reaches saveoptions.
// These aliases keep one spelling of an address across both: the query package
// re-exports CellRef from here, and a value built here is the same type the
// save options parse.
type (
	// CellRef is a zero-based cell coordinate; (Row: 2, Col: 1) is "B3".
	CellRef = refs.CellRef
	// Area is a rectangular cell region with inclusive start and end cells.
	Area = refs.Area
)

// The worksheet grid bounds, re-exported from refs.
const (
	MaxGridRows = refs.MaxGridRows
	MaxGridCols = refs.MaxGridCols
)

// ParseCellRef parses an Excel-style cell reference such as "B3".
func ParseCellRef(s string) (CellRef, error) { return refs.ParseCellRef(s) }

// ParseArea parses an Excel-style area such as "A1:C3"; a single cell is a
// one-cell area.
func ParseArea(s string) (Area, error) { return refs.ParseArea(s) }

// ParseAreaWithinGrid parses an Excel-style area and requires both corners to
// lie inside the worksheet grid.
func ParseAreaWithinGrid(s string) (Area, error) { return refs.ParseAreaWithinGrid(s) }

package query

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// NamedRange is a workbook-level named range: a name bound to a cell area. The
// name lives on the workbook (not on a worksheet), and Area is the primary
// contiguous area it refers to.
type NamedRange struct {
	// Name is the range's name, e.g. "MyRange".
	Name string
	// RefersTo is the raw reference text the engine stores, e.g.
	// "=Data!$A$1:$B$2".
	RefersTo string
	// Area is the referred cell area. It is the zero value when the name is a
	// formula name with no resolvable contiguous range.
	Area Area
}

// NamedRanges lists the workbook's defined names. Options are accepted for API
// uniformity; none currently affect this workbook-level listing.
//
// Example:
//
//	ranges, err := query.NamedRanges(datasource.FilePathSource("data.xlsx"))
//	for _, r := range ranges {
//		fmt.Println(r.Name, r.RefersTo, r.Area)
//	}
func NamedRanges(source datasource.DataSource, opts ...Option) ([]NamedRange, error) {
	applyOptions(defaultOptions(), opts)
	if source == nil {
		return nil, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return nil, err
	}
	names, err := cells.WorkbookNames(workbook)
	if err != nil {
		return nil, err
	}
	count, err := names.GetCount()
	if err != nil {
		return nil, err
	}
	out := make([]NamedRange, 0, count)
	for i := int32(0); i < count; i++ {
		n, err := names.Get_Int(i)
		if err != nil {
			return nil, err
		}
		name, err := n.GetText()
		if err != nil {
			return nil, err
		}
		refersTo, err := n.GetRefersTo()
		if err != nil {
			return nil, err
		}
		nr := NamedRange{Name: name, RefersTo: refersTo}
		if area, ok := nameRangeArea(n); ok {
			nr.Area = area
		}
		out = append(out, nr)
	}
	return out, nil
}

// nameRangeArea resolves a name's primary referred area from its Range object.
// A formula name (e.g. one that refers to a computed value rather than a cell
// block) has no contiguous range; for those the second result is false and the
// caller leaves the area zero.
func nameRangeArea(n *asposecells.Name) (Area, bool) {
	rng, err := n.GetRange()
	if err != nil {
		return Area{}, false
	}
	firstRow, err := rng.GetFirstRow()
	if err != nil {
		return Area{}, false
	}
	firstCol, err := rng.GetFirstColumn()
	if err != nil {
		return Area{}, false
	}
	rowCount, err := rng.GetRowCount()
	if err != nil {
		return Area{}, false
	}
	colCount, err := rng.GetColumnCount()
	if err != nil {
		return Area{}, false
	}
	return Area{
		Start: CellRef{Row: int(firstRow), Col: int(firstCol)},
		End:   CellRef{Row: int(firstRow + rowCount - 1), Col: int(firstCol + colCount - 1)},
	}, true
}

// ReadNamedRange reads the cells of a named range as a row-major grid, the same
// shape as ReadRange. The range may live on any worksheet; the name is resolved
// through the engine's Range object, so this works regardless of the current
// sheet selection. A name that does not exist returns ErrNameNotFound.
//
// Example:
//
//	grid, err := query.ReadNamedRange(datasource.FilePathSource("data.xlsx"), "MyRange")
func ReadNamedRange(source datasource.DataSource, name string, opts ...Option) ([][]CellValue, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return nil, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return nil, err
	}
	n, err := cells.FindName(workbook, name)
	if err != nil {
		return nil, err
	}
	rng, err := n.GetRange()
	if err != nil {
		return nil, err
	}
	ws, err := rng.GetWorksheet()
	if err != nil {
		return nil, err
	}
	firstRow, err := rng.GetFirstRow()
	if err != nil {
		return nil, err
	}
	firstCol, err := rng.GetFirstColumn()
	if err != nil {
		return nil, err
	}
	rowCount, err := rng.GetRowCount()
	if err != nil {
		return nil, err
	}
	colCount, err := rng.GetColumnCount()
	if err != nil {
		return nil, err
	}
	return readGrid(ws, firstRow, firstCol, rowCount, colCount, cfg.trimSpace)
}

// readGrid reads a rectangular block of cells starting at (row, col) with the
// given dimensions, row-major. grid[r][c] is the cell at row row+r, column col+c.
func readGrid(ws *asposecells.Worksheet, row, col, rowCount, colCount int32, trim bool) ([][]CellValue, error) {
	cs, err := ws.GetCells()
	if err != nil {
		return nil, err
	}
	grid := make([][]CellValue, rowCount)
	for r := int32(0); r < rowCount; r++ {
		grid[r] = make([]CellValue, colCount)
		for c := int32(0); c < colCount; c++ {
			cell, err := cs.Get_Int_Int(row+r, col+c)
			if err != nil {
				return nil, err
			}
			v, err := fromCell(cell)
			if err != nil {
				return nil, err
			}
			grid[r][c] = applyTrim(v, trim)
		}
	}
	return grid, nil
}

// ReadCellComment returns the comment note on a single cell, identified by its
// Excel reference, e.g. "B3". A cell without a comment returns an empty string.
//
// Example:
//
//	note, err := query.ReadCellComment(datasource.FilePathSource("data.xlsx"), "A1")
func ReadCellComment(source datasource.DataSource, ref string, opts ...Option) (string, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return "", toolkiterrors.ErrDataSourceNil
	}
	parsed, err := ParseCellRef(ref)
	if err != nil {
		return "", err
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return "", err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return "", err
	}
	return cells.CellComment(ws, int32(parsed.Row), int32(parsed.Col))
}

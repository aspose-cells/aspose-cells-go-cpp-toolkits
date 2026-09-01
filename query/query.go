// Package query reads spreadsheet data into Go values.
//
// It is the read counterpart of transfer: where transfer serializes a sheet or
// range to a format (JSON, XML), query reads cell values into memory so a
// program can inspect and decide on them. Every entry point takes a
// datasource.DataSource plus query Options and returns Go-native data — typed
// CellValue grids, names, dimensions, or merged regions — never engine objects.
//
// By default a query targets the first worksheet by index, because evaluation
// mode occasionally corrupts worksheet names at load time; use WithSheetIndex
// to pick a sheet by position or WithSheet to pick it by name.
package query

import (
	"fmt"
	"strings"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// ReadCell reads a single cell identified by its Excel reference, e.g. "B3".
//
// Example:
//
//	v, err := query.ReadCell(datasource.FilePathSource("data.xlsx"), "B3")
//	if v.Kind() == query.KindText {
//		text, _ := v.String()
//	}
func ReadCell(source datasource.DataSource, ref string, opts ...Option) (CellValue, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return CellValue{}, toolkiterrors.ErrDataSourceNil
	}
	parsed, err := ParseCellRef(ref)
	if err != nil {
		return CellValue{}, err
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return CellValue{}, err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return CellValue{}, err
	}
	cell, err := cells.GetCell(ws, int32(parsed.Row), int32(parsed.Col))
	if err != nil {
		return CellValue{}, err
	}
	v, err := fromCell(cell)
	if err != nil {
		return CellValue{}, err
	}
	return applyTrim(v, cfg.trimSpace), nil
}

// ReadRange reads a rectangular block of cells bounded by startCell and
// endCell, e.g. "A1" and "C3". The result is row-major: grid[r][c] is the
// cell at row start.Row+r, column start.Col+c. A start cell below or to the
// right of the end cell returns ErrInvalidRange.
func ReadRange(source datasource.DataSource, startCell, endCell string, opts ...Option) ([][]CellValue, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return nil, toolkiterrors.ErrDataSourceNil
	}
	start, err := ParseCellRef(startCell)
	if err != nil {
		return nil, err
	}
	end, err := ParseCellRef(endCell)
	if err != nil {
		return nil, err
	}
	if start.Row > end.Row || start.Col > end.Col {
		return nil, fmt.Errorf("invalid range %s:%s: %w", startCell, endCell, toolkiterrors.ErrInvalidRange)
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return nil, err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return nil, err
	}
	rowCount := end.Row - start.Row + 1
	colCount := end.Col - start.Col + 1
	cs, err := ws.GetCells()
	if err != nil {
		return nil, err
	}
	grid := make([][]CellValue, rowCount)
	for r := 0; r < rowCount; r++ {
		grid[r] = make([]CellValue, colCount)
		for c := 0; c < colCount; c++ {
			cell, err := cs.Get_Int_Int(int32(start.Row+r), int32(start.Col+c))
			if err != nil {
				return nil, err
			}
			v, err := fromCell(cell)
			if err != nil {
				return nil, err
			}
			grid[r][c] = applyTrim(v, cfg.trimSpace)
		}
	}
	return grid, nil
}

// ReadWorksheet reads the worksheet's full used range as a row-major grid.
// An empty worksheet yields an empty (length 0) slice.
func ReadWorksheet(source datasource.DataSource, opts ...Option) ([][]CellValue, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return nil, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return nil, err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return nil, err
	}
	rows, cols, err := cells.UsedRange(ws)
	if err != nil {
		return nil, err
	}
	cs, err := ws.GetCells()
	if err != nil {
		return nil, err
	}
	grid := make([][]CellValue, rows)
	for r := int32(0); r < rows; r++ {
		grid[r] = make([]CellValue, cols)
		for c := int32(0); c < cols; c++ {
			cell, err := cs.Get_Int_Int(r, c)
			if err != nil {
				return nil, err
			}
			v, err := fromCell(cell)
			if err != nil {
				return nil, err
			}
			grid[r][c] = applyTrim(v, cfg.trimSpace)
		}
	}
	return grid, nil
}

// ReadMergedCells returns the worksheet's merged cell regions, each reported
// exactly once from its anchor.
func ReadMergedCells(source datasource.DataSource, opts ...Option) ([]Area, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return nil, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return nil, err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return nil, err
	}
	return cells.MergedAreas(ws)
}

// SheetNames returns the names of the workbook's worksheets in order.
func SheetNames(source datasource.DataSource, opts ...Option) ([]string, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return nil, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return nil, err
	}
	return cells.SheetNames(workbook)
}

// Dimensions returns the used range's dimensions (rows, cols). An empty
// worksheet yields (0, 0).
func Dimensions(source datasource.DataSource, opts ...Option) (rows, cols int, err error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return 0, 0, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return 0, 0, err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return 0, 0, err
	}
	r, c, err := cells.UsedRange(ws)
	if err != nil {
		return 0, 0, err
	}
	return int(r), int(c), nil
}

// sheetFor resolves the options' sheet selection: by index when WithSheetIndex
// was used (or the default first-sheet selection), otherwise by name.
func sheetFor(cfg *options, wb *asposecells.Workbook) (*asposecells.Worksheet, error) {
	if cfg.useIndex {
		return cells.SheetByIndex(wb, cfg.sheetIndex)
	}
	wss, err := wb.GetWorksheets()
	if err != nil {
		return nil, err
	}
	return cells.WorksheetByName(wss, cfg.sheetName)
}

// applyTrim trims a text value when the WithTrimSpace option requests it.
func applyTrim(v CellValue, trim bool) CellValue {
	if trim && v.kind == KindText {
		v.s = strings.TrimSpace(v.s)
	}
	return v
}

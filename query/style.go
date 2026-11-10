package query

import (
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
)

// Style describes the formatting of a single cell in toolkit-native values.
// It is a snapshot: reading a Style does not keep any engine object alive.
type Style struct {
	// FontName is the font family name, e.g. "Calibri".
	FontName string
	// FontSize is the font size in points.
	FontSize int32
	// Bold and Italic report the font's weight and slant.
	Bold   bool
	Italic bool
	// FontColor is the font color as an "#RRGGBB" hex string.
	FontColor string
	// NumberFormat is the built-in number format id (0 = General).
	NumberFormat int32
	// CustomFormat is the custom number format string, empty for a built-in
	// format.
	CustomFormat string
}

// ReadCellFormula reads the formula of a single cell identified by its Excel
// reference, e.g. "B3". A cell that holds a value rather than a formula returns
// an empty string.
//
// Example:
//
//	formula, err := query.ReadCellFormula(
//		datasource.FilePathSource("data.xlsx"), "C3")
func ReadCellFormula(source datasource.DataSource, ref string, opts ...Option) (string, error) {
	engine.LockEngine()
	defer engine.UnlockEngine()
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return "", toolkiterrors.ErrDataSourceNil
	}
	parsed, err := ParseCellRefWithinGrid(ref)
	if err != nil {
		return "", err
	}
	workbook, err := cells.GetWorkbookWithDataSourceAndOptions(source, cfg.load)
	if err != nil {
		return "", err
	}
	defer engine.CloseWorkbook(workbook)
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return "", err
	}
	cell, err := cells.GetCell(ws, int32(parsed.Row), int32(parsed.Col))
	if err != nil {
		return "", err
	}
	return cell.GetFormula()
}

// ReadStyle reads the formatting of a single cell identified by its Excel
// reference, e.g. "B3". The returned Style holds plain Go values, so it remains
// valid after the workbook has been closed.
//
// Example:
//
//	style, err := query.ReadStyle(
//		datasource.FilePathSource("data.xlsx"), "B3")
//	if style.Bold {
//		// ...
//	}
func ReadStyle(source datasource.DataSource, ref string, opts ...Option) (Style, error) {
	engine.LockEngine()
	defer engine.UnlockEngine()
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return Style{}, toolkiterrors.ErrDataSourceNil
	}
	parsed, err := ParseCellRefWithinGrid(ref)
	if err != nil {
		return Style{}, err
	}
	workbook, err := cells.GetWorkbookWithDataSourceAndOptions(source, cfg.load)
	if err != nil {
		return Style{}, err
	}
	defer engine.CloseWorkbook(workbook)
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return Style{}, err
	}
	cell, err := cells.GetCell(ws, int32(parsed.Row), int32(parsed.Col))
	if err != nil {
		return Style{}, err
	}
	snapshot, err := cells.ReadCellStyle(cell)
	if err != nil {
		return Style{}, err
	}
	return Style{
		FontName:     snapshot.FontName,
		FontSize:     snapshot.FontSize,
		Bold:         snapshot.Bold,
		Italic:       snapshot.Italic,
		FontColor:    snapshot.FontColor,
		NumberFormat: snapshot.NumberFormat,
		CustomFormat: snapshot.CustomFormat,
	}, nil
}

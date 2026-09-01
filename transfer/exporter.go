// Package transfer exports spreadsheet data to structured formats and imports
// structured data back into spreadsheets.
//
// Exports cover XML and per-worksheet / per-range JSON; imports cover CSV, XML,
// and JSON data into a worksheet. Every entry point writes its result to a
// datasource.DataSink, so the caller picks the output shape (file, writer, or
// in-memory bytes) by choosing the sink, and configures the target sheet, cell
// area, and format-specific options with transfer Options.
package transfer

import (
	"fmt"
	"os"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	jsonsaveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/json"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Option configures a transfer entry point.
type Option func(*options)

type options struct {
	sheet          cells.Sheet
	startCell      string
	endCell        string
	xmlMap         string
	beginRow       int
	beginColumn    int
	convertNumeric bool
	separator      string
}

func defaultOptions() *options {
	// Default to the first worksheet by index (matching query), so transfer
	// shares query's immunity to evaluation-mode name corruption.
	return &options{
		sheet:          cells.FirstSheet,
		startCell:      "A1",
		endCell:        "",
		xmlMap:         "Sheet1",
		convertNumeric: true,
		separator:      ",",
	}
}

// WithSheet sets the worksheet to export from or import into, by name.
//
// Note: in evaluation mode the engine occasionally corrupts a worksheet's name
// when the workbook is loaded (observed ~2% of loads, any sheet, not just the
// default first sheet), so a name-based lookup may fail with
// ErrWorksheetNotFound even for a sheet that exists. Prefer WithSheetIndex,
// which is immune to name corruption. When targeting by name, use an
// explicitly named sheet — create it via editor.WithAddWorksheet or rename the
// source sheet — and callers that cannot tolerate the occasional spurious miss
// should retry the whole operation on fresh input.
func WithSheet(name string) Option {
	return func(o *options) { o.sheet = cells.Sheet{Name: name} }
}

// WithSheetIndex sets the worksheet to export from or import into by its
// zero-based index. Index-based lookup is immune to the evaluation-mode
// load-time name corruption, so it is the recommended way to target a sheet.
// The default targets the first worksheet by index.
func WithSheetIndex(i int) Option {
	return func(o *options) { o.sheet = cells.Sheet{UseIndex: true, Index: i} }
}

// WithStartCell sets the top-left cell of an export range, e.g. "A1".
func WithStartCell(ref string) Option {
	return func(o *options) { o.startCell = ref }
}

// WithEndCell sets the bottom-right cell of an export range, e.g. "B3". When
// unset, the range extends to the worksheet's last used cell.
func WithEndCell(ref string) Option {
	return func(o *options) { o.endCell = ref }
}

// WithXMLMap sets the XML map name used by ExportSpreadsheetToXml, e.g.
// "InventoryMap".
func WithXMLMap(name string) Option {
	return func(o *options) { o.xmlMap = name }
}

// WithBeginCell sets the top-left cell of an import as a (row, column) pair,
// both zero-based.
func WithBeginCell(row, col int) Option {
	return func(o *options) { o.beginRow, o.beginColumn = row, col }
}

// WithConvertNumeric controls whether numeric-looking CSV fields are imported
// as numbers rather than text. Defaults to true.
func WithConvertNumeric(b bool) Option {
	return func(o *options) { o.convertNumeric = b }
}

// WithSeparator sets the field separator used by ImportCSV. Defaults to ",".
func WithSeparator(s string) Option {
	return func(o *options) { o.separator = s }
}

func applyOptions(o *options, opts []Option) {
	for _, opt := range opts {
		if opt != nil {
			opt(o)
		}
	}
}

// ExportWorksheetToJson exports a worksheet's used range as JSON. The result is
// written to sink; by default the whole used area of sheet "Sheet1" is
// exported, and WithSheet / WithStartCell / WithEndCell narrow the target.
//
// Example:
//
//	err := transfer.ExportWorksheetToJson(
//		datasource.FilePathSource("out/seed.xlsx"),
//		datasource.FilePathSink("out/sheet.json"),
//		transfer.WithSheet("Data"))
func ExportWorksheetToJson(source datasource.DataSource, sink datasource.DataSink, opts ...Option) error {
	return ExportRangeToJson(source, sink, opts...)
}

// ExportRangeToJson exports a cell range as JSON. The range defaults to the
// worksheet's full used area (WithStartCell "A1", no end cell); use
// WithStartCell and WithEndCell to export a specific range.
//
// Example:
//
//	err := transfer.ExportRangeToJson(
//		datasource.FilePathSource("out/seed.xlsx"),
//		datasource.FilePathSink("out/range.json"),
//		transfer.WithSheet("Data"), transfer.WithStartCell("A1"),
//		transfer.WithEndCell("B3"))
func ExportRangeToJson(source datasource.DataSource, sink datasource.DataSink, opts ...Option) error {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return toolkiterrors.ErrDataSourceNil
	}
	if sink == nil {
		return toolkiterrors.ErrDataSinkNil
	}
	data, err := cells.ReadSource(source)
	if err != nil {
		return err
	}
	workbook, err := asposecells.NewWorkbook_Stream(data)
	if err != nil {
		return err
	}
	ws, err := cfg.sheet.Resolve(workbook)
	if err != nil {
		return err
	}
	sheetIndex, err := ws.GetIndex()
	if err != nil {
		return err
	}
	startRow, startColumn, err := asposecells.CellsHelper_CellNameToIndex(cfg.startCell)
	if err != nil {
		return err
	}
	var endRow, endColumn int32
	if cfg.endCell == "" {
		// Full used area: extend the range to the last used cell.
		worksheetCells, err := ws.GetCells()
		if err != nil {
			return err
		}
		if endRow, err = worksheetCells.GetMaxDataRow(); err != nil {
			return err
		}
		if endColumn, err = worksheetCells.GetMaxDataColumn(); err != nil {
			return err
		}
	} else {
		if endRow, endColumn, err = asposecells.CellsHelper_CellNameToIndex(cfg.endCell); err != nil {
			return err
		}
	}
	cellArea, err := asposecells.CellArea_CreateCellArea_Int_Int_Int_Int(startRow, startColumn, endRow, endColumn)
	if err != nil {
		return err
	}
	opt := jsonsaveoptions.New(jsonsaveoptions.WithSheetIndexes([]int32{sheetIndex}), jsonsaveoptions.WithExportArea(cellArea))
	out, err := opt.Apply(data)
	if err != nil {
		return err
	}
	// An empty range (e.g. a worksheet with no data) would otherwise produce a
	// successful 0-byte result, which callers can't distinguish from a missing
	// file. Write a valid empty JSON array instead.
	if len(out) == 0 {
		out = []byte("[]")
	}
	return sink.Write("", out)
}

// ExportSpreadsheetToXml exports the whole workbook as XML using the given XML
// map name (WithXMLMap), writing the result to sink.
//
// Example:
//
//	err := transfer.ExportSpreadsheetToXml(
//		datasource.FilePathSource("out/seed.xlsx"),
//		datasource.FilePathSink("out/inventory.xml"),
//		transfer.WithXMLMap("InventoryMap"))
func ExportSpreadsheetToXml(source datasource.DataSource, sink datasource.DataSink, opts ...Option) error {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return toolkiterrors.ErrDataSourceNil
	}
	if sink == nil {
		return toolkiterrors.ErrDataSinkNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return err
	}
	data, err := workbook.ExportXml_String(cfg.xmlMap)
	if err != nil {
		return err
	}
	return sink.Write("", data)
}

// ExportWorksheetToJsonFile exports a worksheet's used range as JSON straight
// from file to file.
//
// Deprecated: use ExportWorksheetToJson with a datasource.FilePathSource and
// datasource.FilePathSink instead.
func ExportWorksheetToJsonFile(spreadsheet string, worksheet string, outputPath string) error {
	if err := requireFile(spreadsheet); err != nil {
		return err
	}
	return ExportWorksheetToJson(datasource.FilePathSource(spreadsheet), datasource.FilePathSink(outputPath), WithSheet(worksheet))
}

// ExportRangeToJsonFile exports a cell range as JSON straight from file to
// file.
//
// Deprecated: use ExportRangeToJson with a datasource.FilePathSource and
// datasource.FilePathSink instead.
func ExportRangeToJsonFile(spreadsheet string, worksheet string, startCellName string, endCellName string, outputPath string) error {
	if err := requireFile(spreadsheet); err != nil {
		return err
	}
	return ExportRangeToJson(
		datasource.FilePathSource(spreadsheet),
		datasource.FilePathSink(outputPath),
		WithSheet(worksheet), WithStartCell(startCellName), WithEndCell(endCellName),
	)
}

// ExportSpreadsheetToXmlFile exports the whole workbook as XML straight from
// file to file.
//
// Deprecated: use ExportSpreadsheetToXml with a datasource.FilePathSource and
// datasource.FilePathSink instead.
func ExportSpreadsheetToXmlFile(spreadsheet string, mapName string, outputPath string) error {
	if err := requireFile(spreadsheet); err != nil {
		return err
	}
	return ExportSpreadsheetToXml(datasource.FilePathSource(spreadsheet), datasource.FilePathSink(outputPath), WithXMLMap(mapName))
}

// requireFile returns ErrInputIsFolder when the given path is a directory.
func requireFile(path string) error {
	fileInfo, err := os.Stat(path)
	if err != nil {
		return err
	}
	if fileInfo.IsDir() {
		return fmt.Errorf("%q is a folder, expected a file: %w", path, toolkiterrors.ErrInputIsFolder)
	}
	return nil
}

package editor

import (
	"fmt"
	"reflect"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	rowmap "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/rows"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// WriteRowsOption configures WriteRows.
type WriteRowsOption func(*writeRowsConfig)

type writeRowsConfig struct {
	sheet       cells.Sheet
	writeHeader bool
}

func defaultWriteRowsConfig() *writeRowsConfig {
	// Default to the first worksheet by index, aligned with query and transfer.
	return &writeRowsConfig{sheet: cells.FirstSheet}
}

func applyWriteRowsOptions(cfg *writeRowsConfig, opts []WriteRowsOption) {
	for _, opt := range opts {
		if opt != nil {
			opt(cfg)
		}
	}
}

// WithWriteHeader controls whether WriteRows writes a header row of column
// names before the data rows. Defaults to false.
func WithWriteHeader(b bool) WriteRowsOption {
	return func(c *writeRowsConfig) { c.writeHeader = b }
}

// WithSheet selects the worksheet to write to by name.
//
// Note: in evaluation mode the engine occasionally corrupts a worksheet's name
// when the workbook is loaded (observed ~2% of loads), so a name-based lookup
// may fail with ErrWorksheetNotFound even for a sheet that exists. Prefer
// WithSheetIndex, which is immune to name corruption.
func WithSheet(name string) WriteRowsOption {
	return func(c *writeRowsConfig) { c.sheet = cells.Sheet{Name: name} }
}

// WithSheetIndex selects the worksheet to write to by its zero-based index.
// Index-based lookup is immune to the evaluation-mode load-time name
// corruption, so it is the recommended way to target a sheet.
func WithSheetIndex(i int) WriteRowsOption {
	return func(c *writeRowsConfig) { c.sheet = cells.Sheet{UseIndex: true, Index: i} }
}

// WriteRows writes a slice of T into a worksheet and saves the result to a
// sink. It is the write counterpart of query.ReadRows and uses the same column
// mapping: each struct field becomes a column named by its `excel:"name"` tag,
// or by its own name when no tag is present; `excel:"-"` skips a field. With
// WriteRows with WithWriteHeader(true) writes those column names as a header
// row (row 0) before the data rows; otherwise data starts at row 0.
//
// Supported field values are the types SetCellValue accepts: integers,
// unsigned integers, floats, string, bool, and time.Time. The workbook is
// saved in its original format. Cells outside the written block are left
// untouched.
//
// Example:
//
//	type Employee struct {
//		ID   int    `excel:"id"`
//		Name string `excel:"name"`
//	}
//	err := editor.WriteRows(
//		datasource.FilePathSource("template.xlsx"),
//		datasource.FilePathSink("employees.xlsx"),
//		[]Employee{{ID: 1, Name: "Ada"}},
//		editor.WithSheetIndex(0),
//		editor.WithWriteHeader(true),
//	)
func WriteRows[T any](source datasource.DataSource, sink datasource.DataSink, rows []T, opts ...WriteRowsOption) error {
	cfg := defaultWriteRowsConfig()
	applyWriteRowsOptions(cfg, opts)
	if source == nil {
		return toolkiterrors.ErrDataSourceNil
	}
	if sink == nil {
		return toolkiterrors.ErrDataSinkNil
	}
	var zero T
	typ := reflect.TypeOf(zero)
	if typ.Kind() != reflect.Struct {
		return fmt.Errorf("WriteRows: T must be a struct, got %s: %w", typ, toolkiterrors.ErrInvalidValue)
	}
	cols, err := rowmap.Columns(typ)
	if err != nil {
		return err
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return err
	}
	ws, err := cfg.sheet.Resolve(workbook)
	if err != nil {
		return err
	}
	if err := writeRows(ws, cols, rows, cfg.writeHeader); err != nil {
		return err
	}
	data, err := cells.WorkbookToByteData(workbook)
	if err != nil {
		return err
	}
	return sink.Write("", data)
}

// writeRows writes the optional header row and each row's fields through
// SetCellValue so the value-type handling (including the date number format
// applied to time.Time) stays in one place.
func writeRows[T any](ws *asposecells.Worksheet, cols []rowmap.Column, rows []T, writeHeader bool) error {
	start := 0
	if writeHeader {
		for i, c := range cols {
			if err := SetCellValue(0, i, c.Name)(ws); err != nil {
				return fmt.Errorf("header column %q: %w", c.Name, err)
			}
		}
		start = 1
	}
	for r, item := range rows {
		rv := reflect.ValueOf(item)
		for i, c := range cols {
			if err := SetCellValue(start+r, i, rv.Field(c.Index).Interface())(ws); err != nil {
				return fmt.Errorf("row %d, column %q: %w", r+1, c.Name, err)
			}
		}
	}
	return nil
}

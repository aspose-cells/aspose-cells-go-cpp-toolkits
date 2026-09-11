package tests

import (
	"errors"
	"fmt"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/transfer"
)

// TestTransferImportJsonData verifies ImportJsonData actually lands the JSON in
// the target worksheet, not merely that the workbook round-trips: the engine's
// JSON layout writes the outer key as a header, the object keys as the next
// header row, then one row per object, with JSON numbers arriving as numeric
// cells. The read-back runs under retryStable because evaluation mode can
// corrupt a loaded cell's value.
func TestTransferImportJsonData(t *testing.T) {
	// Create a seed workbook
	seedBytes := newTestWorkbookBytes(t)

	// JSON data to import
	jsonData := []byte(`{
		"values": [
			{"name": "Alice", "age": 30},
			{"name": "Bob", "age": 25}
		]
	}`)

	// Import JSON data
	var out datasource.BytesSink
	err := transfer.ImportJsonData(
		datasource.BytesSource(seedBytes),
		datasource.BytesSource(jsonData),
		&out,
		// Target the first sheet by index: transfer.WithSheet resolves by name,
		// and evaluation mode (no license) rewrites a random sheet's name on
		// ~2% of loads, which would flake with ErrWorksheetNotFound.
		transfer.WithSheetIndex(0),
		transfer.WithBeginCell(0, 0),
	)
	if err != nil {
		t.Fatalf("ImportJsonData error: %v", err)
	}

	result := out.Bytes()
	if len(result) == 0 {
		t.Fatal("ImportJsonData returned empty result")
	}

	if err := retryStable(5, func() error {
		return verifyImportedGrid(datasource.BytesSource(result), [][]interface{}{
			{"values", nil},
			{"name", "age"},
			{"Alice", 30},
			{"Bob", 25},
		})
	}); err != nil {
		t.Fatal(err)
	}
}

// TestTransferImportXmlData verifies ImportXMLData actually lands the XML in the
// target worksheet: the engine writes one column per distinct child element
// (header row first) and one row per <row>, with every value read back as text.
// The read-back runs under retryStable because evaluation mode can corrupt a
// loaded cell's value.
func TestTransferImportXmlData(t *testing.T) {
	// Create a seed workbook
	seedBytes := newTestWorkbookBytes(t)

	// XML data to import
	xmlData := []byte(`<?xml version="1.0" encoding="utf-8"?>
<root>
	<row>
		<name>Alice</name>
		<age>30</age>
	</row>
	<row>
		<name>Bob</name>
		<age>25</age>
	</row>
</root>`)

	// Import XML data
	var out datasource.BytesSink
	err := transfer.ImportXMLData(
		datasource.BytesSource(seedBytes),
		datasource.BytesSource(xmlData),
		&out,
		// See TestTransferImportJsonData: index-based selection is immune to
		// the evaluation-mode load-time sheet-name corruption.
		transfer.WithSheetIndex(0),
		transfer.WithBeginCell(0, 0),
	)
	if err != nil {
		t.Fatalf("ImportXMLData error: %v", err)
	}

	result := out.Bytes()
	if len(result) == 0 {
		t.Fatal("ImportXMLData returned empty result")
	}

	if err := retryStable(5, func() error {
		return verifyImportedGrid(datasource.BytesSource(result), [][]interface{}{
			{"name", "age"},
			{"Alice", "30"},
			{"Bob", "25"},
		})
	}); err != nil {
		t.Fatal(err)
	}
}

// verifyImportedGrid loads src and compares its used range against want cell by
// cell, so an import that silently no-ops fails instead of passing on
// "the output was non-empty". A nil entry expects an empty cell; string and int
// entries assert the cell's kind as well as its value. It returns an error so
// callers can re-load under retryStable.
func verifyImportedGrid(src datasource.DataSource, want [][]interface{}) error {
	grid, err := query.ReadWorksheet(src)
	if err != nil {
		return fmt.Errorf("ReadWorksheet: %w", err)
	}
	if len(grid) != len(want) {
		return fmt.Errorf("imported grid has %d rows, want %d", len(grid), len(want))
	}
	for r, wantRow := range want {
		if len(grid[r]) != len(wantRow) {
			return fmt.Errorf("row %d has %d columns, want %d", r, len(grid[r]), len(wantRow))
		}
		for c, w := range wantRow {
			got := grid[r][c]
			text, _ := got.String()
			switch w := w.(type) {
			case nil:
				if got.Kind() != query.KindEmpty {
					return fmt.Errorf("cell(%d,%d) = %s %q, want empty", r, c, got.Kind(), text)
				}
			case string:
				if got.Kind() != query.KindText || text != w {
					return fmt.Errorf("cell(%d,%d) = %s %q, want text %q", r, c, got.Kind(), text, w)
				}
			case int:
				i, ok := got.Int()
				if got.Kind() != query.KindInt || !ok || i != int64(w) {
					return fmt.Errorf("cell(%d,%d) = %s %q, want int %d", r, c, got.Kind(), text, w)
				}
			default:
				return fmt.Errorf("cell(%d,%d): unsupported expectation %T", r, c, w)
			}
		}
	}
	return nil
}

// TestTransferImportJsonDataNilSource verifies ErrDataSourceNil is returned when source is nil.
func TestTransferImportJsonDataNilSource(t *testing.T) {
	var out datasource.BytesSink
	err := transfer.ImportJsonData(nil, datasource.BytesSource([]byte{}), &out)
	if !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Fatalf("ImportJsonData(nil, ...) error = %v, want ErrDataSourceNil", err)
	}
}

// TestTransferImportJsonDataNilJsonData verifies ErrDataSourceNil is returned when jsonData is nil.
func TestTransferImportJsonDataNilJsonData(t *testing.T) {
	seedBytes := newTestWorkbookBytes(t)
	var out datasource.BytesSink
	err := transfer.ImportJsonData(datasource.BytesSource(seedBytes), nil, &out)
	if !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Fatalf("ImportJsonData(..., nil) error = %v, want ErrDataSourceNil", err)
	}
}

// TestTransferImportJsonDataNilSink verifies ErrDataSinkNil is returned when sink is nil.
func TestTransferImportJsonDataNilSink(t *testing.T) {
	seedBytes := newTestWorkbookBytes(t)
	err := transfer.ImportJsonData(datasource.BytesSource(seedBytes), datasource.BytesSource([]byte{}), nil)
	if !errors.Is(err, toolkiterrors.ErrDataSinkNil) {
		t.Fatalf("ImportJsonData(..., nil) error = %v, want ErrDataSinkNil", err)
	}
}

// TestTransferImportXmlDataNilSource verifies ErrDataSourceNil is returned when source is nil.
func TestTransferImportXmlDataNilSource(t *testing.T) {
	var out datasource.BytesSink
	err := transfer.ImportXMLData(nil, datasource.BytesSource([]byte{}), &out)
	if !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Fatalf("ImportXMLData(nil, ...) error = %v, want ErrDataSourceNil", err)
	}
}

// TestTransferImportXmlDataNilXmlData verifies ErrDataSourceNil is returned when xmlData is nil.
func TestTransferImportXmlDataNilXmlData(t *testing.T) {
	seedBytes := newTestWorkbookBytes(t)
	var out datasource.BytesSink
	err := transfer.ImportXMLData(datasource.BytesSource(seedBytes), nil, &out)
	if !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Fatalf("ImportXMLData(..., nil) error = %v, want ErrDataSourceNil", err)
	}
}

// TestTransferImportXmlDataNilSink verifies ErrDataSinkNil is returned when sink is nil.
func TestTransferImportXmlDataNilSink(t *testing.T) {
	seedBytes := newTestWorkbookBytes(t)
	err := transfer.ImportXMLData(datasource.BytesSource(seedBytes), datasource.BytesSource([]byte{}), nil)
	if !errors.Is(err, toolkiterrors.ErrDataSinkNil) {
		t.Fatalf("ImportXMLData(..., nil) error = %v, want ErrDataSinkNil", err)
	}
}

package tests

import (
	"errors"
	"fmt"
	"strings"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/transfer"
)

// TestParseAreaWithinGrid pins the grid check that guards the A1-area option
// values. Parsing alone is not enough: the parser accepts any row number and any
// run of column letters, and the engine accepts an off-grid area and then
// matches nothing, so without this check a typo is a silently empty export.
func TestParseAreaWithinGrid(t *testing.T) {
	valid := []struct {
		in       string
		startRow int
		startCol int
		endRow   int
		endCol   int
	}{
		{"A1:C3", 0, 0, 2, 2},
		{"B2", 1, 1, 1, 1},                             // a single cell is a one-cell area
		{"XFD1048576", 1048575, 16383, 1048575, 16383}, // the far corner
		{"a1:c3", 0, 0, 2, 2},                          // column letters are case-insensitive
	}
	for _, tc := range valid {
		area, err := cells.ParseAreaWithinGrid(tc.in)
		if err != nil {
			t.Errorf("ParseAreaWithinGrid(%q) error: %v", tc.in, err)
			continue
		}
		if area.Start.Row != tc.startRow || area.Start.Col != tc.startCol ||
			area.End.Row != tc.endRow || area.End.Col != tc.endCol {
			t.Errorf("ParseAreaWithinGrid(%q) = %+v, want [%d,%d .. %d,%d]",
				tc.in, area, tc.startRow, tc.startCol, tc.endRow, tc.endCol)
		}
	}

	invalid := []struct {
		in   string
		want error
	}{
		{"XFE1:XFE9", toolkiterrors.ErrInvalidRange},          // column past XFD
		{"A1:A1048577", toolkiterrors.ErrInvalidRange},        // row past the last one
		{"ZZZZZZZZZZZZZZ1:Z9", toolkiterrors.ErrInvalidRange}, // overflows the column counter
		{"C3:A1", toolkiterrors.ErrInvalidRange},              // backwards
		{"garbage", toolkiterrors.ErrInvalidCellRef},
		{"", toolkiterrors.ErrInvalidCellRef},
	}
	for _, tc := range invalid {
		if _, err := cells.ParseAreaWithinGrid(tc.in); !errors.Is(err, tc.want) {
			t.Errorf("ParseAreaWithinGrid(%q) error = %v, want %v", tc.in, err, tc.want)
		}
	}
}

// TestExportRangeToJsonHonoursArea checks the area actually reaches the engine:
// exporting A1:B2 of a three-row table must not contain the third row. Asserting
// on the JSON text is the point — an option that is accepted and then ignored
// produces a valid file with the wrong contents, which no error surfaces.
func TestExportRangeToJsonHonoursArea(t *testing.T) {
	if err := retryStable(5, func() error {
		seed := datasource.BytesSource(newNamedWorkbookBytes(t, "Data"))
		csv := datasource.BytesSource([]byte("id,name\n1,Alpha\n2,Beta\n"))

		var imported datasource.BytesSink
		if err := transfer.ImportCSV(seed, csv, &imported,
			transfer.WithSheet("Data"), transfer.WithBeginCell(0, 0),
			transfer.WithConvertNumeric(true), transfer.WithSeparator(","),
		); err != nil {
			return fmt.Errorf("ImportCSV: %w", err)
		}

		var out datasource.BytesSink
		if err := transfer.ExportRangeToJson(datasource.BytesSource(imported.Bytes()), &out,
			transfer.WithSheet("Data"), transfer.WithStartCell("A1"), transfer.WithEndCell("B2"),
		); err != nil {
			return fmt.Errorf("ExportRangeToJson: %w", err)
		}
		got := string(out.Bytes())
		if !strings.Contains(got, "Alpha") {
			return fmt.Errorf("A1:B2 export is missing the row inside the area:\n%s", got)
		}
		if strings.Contains(got, "Beta") {
			return fmt.Errorf("A1:B2 export leaked the row outside the area:\n%s", got)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

// TestExportRangeToJsonStartPastUsedRange checks that a start cell beyond the
// worksheet's data yields nothing rather than the whole sheet. The end cell is
// implicit (the last used cell) in this case, so the resolved end falls before
// the start: an implementation that read that as "no area to set" would fall
// back to exporting the entire used range, answering a range request with data
// from outside it.
func TestExportRangeToJsonStartPastUsedRange(t *testing.T) {
	if err := retryStable(5, func() error {
		seed := datasource.BytesSource(newNamedWorkbookBytes(t, "Data"))
		csv := datasource.BytesSource([]byte("id,name\n1,Alpha\n2,Beta\n"))

		var imported datasource.BytesSink
		if err := transfer.ImportCSV(seed, csv, &imported,
			transfer.WithSheet("Data"), transfer.WithBeginCell(0, 0),
			transfer.WithConvertNumeric(true), transfer.WithSeparator(","),
		); err != nil {
			return fmt.Errorf("ImportCSV: %w", err)
		}

		var out datasource.BytesSink
		if err := transfer.ExportRangeToJson(datasource.BytesSource(imported.Bytes()), &out,
			transfer.WithSheet("Data"), transfer.WithStartCell("Z100"),
		); err != nil {
			return fmt.Errorf("ExportRangeToJson: %w", err)
		}
		if got := strings.TrimSpace(string(out.Bytes())); got != "[]" {
			return fmt.Errorf("start past the used range exported %s, want []", got)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

// TestExportRangeToJsonRejectsBackwardsArea checks that a range the caller named
// the wrong way round is reported rather than exported as nothing.
func TestExportRangeToJsonRejectsBackwardsArea(t *testing.T) {
	src := datasource.BytesSource(newNamedWorkbookBytes(t, "Data"))
	var out datasource.BytesSink
	err := transfer.ExportRangeToJson(src, &out,
		transfer.WithSheet("Data"), transfer.WithStartCell("C3"), transfer.WithEndCell("A1"),
	)
	if !errors.Is(err, toolkiterrors.ErrInvalidRange) {
		t.Fatalf("ExportRangeToJson(C3:A1) error = %v, want ErrInvalidRange", err)
	}
}

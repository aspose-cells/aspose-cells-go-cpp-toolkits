package tests

import (
	"errors"
	"fmt"
	"testing"
	"time"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
)

// person is the shared round-trip row type for TestReadRowsRoundTrip. The
// Score field deliberately mixes whole-number and fractional inputs so the
// reader must accept a KindInt cell into a float64 field as well as KindFloat.
type person struct {
	ID     int       `excel:"id"`
	Name   string    `excel:"name"`
	Score  float64   `excel:"score"`
	Active bool      `excel:"active"`
	Joined time.Time `excel:"joined"`
}

// TestReadRowsRoundTrip writes a typed slice through editor.WriteRows, then
// reads it back with query.ReadRows and verifies every field decodes to the
// original value — including a time.Time date and an integer-valued cell read
// into a float64 field. The read-back runs under retryStable because
// evaluation mode can corrupt a random cell's value at load time.
func TestReadRowsRoundTrip(t *testing.T) {
	want := []person{
		{ID: 1, Name: "Ada", Score: 3.14, Active: true, Joined: time.Date(2024, 1, 2, 0, 0, 0, 0, time.UTC)},
		{ID: 2, Name: "Bob", Score: 2, Active: false, Joined: time.Date(2023, 6, 15, 0, 0, 0, 0, time.UTC)},
	}
	var sink datasource.BytesSink
	if err := editor.WriteRows(
		datasource.BytesSource(newTestWorkbookBytes(t)),
		&sink,
		want,
		editor.WithSheetIndex(0),
		editor.WithWriteHeader(true),
	); err != nil {
		t.Fatalf("WriteRows: %v", err)
	}

	src := datasource.BytesSource(sink.Bytes())
	if err := retryStable(2, func() error {
		got, err := query.ReadRows[person](src, query.WithSheetIndex(0))
		if err != nil {
			return fmt.Errorf("ReadRows: %w", err)
		}
		if len(got) != len(want) {
			return fmt.Errorf("ReadRows returned %d rows, want %d", len(got), len(want))
		}
		for i := range want {
			// Compare field-by-field; the date is compared on its wall clock
			// because the engine does not preserve the original time zone.
			if got[i].ID != want[i].ID || got[i].Name != want[i].Name ||
				got[i].Score != want[i].Score || got[i].Active != want[i].Active {
				return fmt.Errorf("row %d = %+v, want %+v", i, got[i], want[i])
			}
			g, w := got[i].Joined, want[i].Joined
			if g.Year() != w.Year() || g.Month() != w.Month() || g.Day() != w.Day() {
				return fmt.Errorf("row %d Joined = %v, want %v", i, g, w)
			}
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

// mapped is the row type for TestReadRowsHeaderMapping: columns are matched by
// name regardless of order, an untagged field matches its own field name, and
// a `excel:"-"` field is skipped entirely.
type mapped struct {
	ID     int    `excel:"id"`
	Name   string `excel:"name"`
	Email  string // no tag: matches the "email" header column by field name
	Secret string `excel:"-"`
}

// TestReadRowsHeaderMapping builds a sheet whose header columns are deliberately
// out of struct-field order and carries an extra "secret" column, then verifies
// ReadRows maps columns by their excel tag (or field name) and skips tagged-out
// fields.
func TestReadRowsHeaderMapping(t *testing.T) {
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(newTestWorkbookBytes(t)),
		editor.InWorksheet(0,
			editor.SetCellValue(0, 0, "email"),
			editor.SetCellValue(0, 1, "id"),
			editor.SetCellValue(0, 2, "name"),
			editor.SetCellValue(0, 3, "secret"),
			editor.SetCellValue(1, 0, "a@x.com"),
			editor.SetCellValue(1, 1, 7),
			editor.SetCellValue(1, 2, "Zoe"),
			editor.SetCellValue(1, 3, "hidden"),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}
	src := datasource.BytesSource(out)
	if err := retryStable(2, func() error {
		got, err := query.ReadRows[mapped](src, query.WithSheetIndex(0))
		if err != nil {
			return fmt.Errorf("ReadRows: %w", err)
		}
		want := []mapped{{ID: 7, Name: "Zoe", Email: "a@x.com", Secret: ""}}
		if len(got) != 1 || got[0] != want[0] {
			return fmt.Errorf("ReadRows = %+v, want %+v", got, want)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

// TestReadRowsEmptySheet verifies an empty worksheet yields an empty (non-nil)
// slice rather than an error.
func TestReadRowsEmptySheet(t *testing.T) {
	rows, err := query.ReadRows[person](datasource.BytesSource(newTestWorkbookBytes(t)))
	if err != nil {
		t.Fatalf("ReadRows on empty sheet: %v", err)
	}
	if rows == nil || len(rows) != 0 {
		t.Errorf("ReadRows on empty sheet = %#v, want empty non-nil slice", rows)
	}
}

// TestReadRowsErrors verifies the reader's error classification: nil source,
// non-struct row type, a header column missing from the struct, and a cell
// whose kind does not fit the field type.
func TestReadRowsErrors(t *testing.T) {
	// T is validated before any load, so a nil source and a non-struct row type
	// fail without touching the engine.
	if _, err := query.ReadRows[person](nil); !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Errorf("ReadRows(nil) error = %v, want ErrDataSourceNil", err)
	}
	if _, err := query.ReadRows[int](datasource.BytesSource([]byte("not a workbook"))); !errors.Is(err, toolkiterrors.ErrInvalidValue) {
		t.Errorf("ReadRows[int] error = %v, want ErrInvalidValue", err)
	}

	// One workbook exercises both a value whose kind does not fit the field
	// type and a header column missing from the struct.
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(newTestWorkbookBytes(t)),
		editor.InWorksheet(0,
			editor.SetCellValue(0, 0, "count"),
			editor.SetCellValue(0, 1, "nope"),
			editor.SetCellValue(1, 0, "not-a-number"),
			editor.SetCellValue(1, 1, "x"),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}
	src := datasource.BytesSource(out)

	type countRow struct {
		Count int `excel:"count"`
	}
	if _, err := query.ReadRows[countRow](src); !errors.Is(err, toolkiterrors.ErrInvalidValue) {
		t.Errorf("ReadRows type mismatch error = %v, want ErrInvalidValue", err)
	}

	type absentCol struct {
		Nope string `excel:"absent"`
	}
	if _, err := query.ReadRows[absentCol](src); !errors.Is(err, toolkiterrors.ErrColumnNotFound) {
		t.Errorf("ReadRows with missing column error = %v, want ErrColumnNotFound", err)
	}
}

// TestWriteRowsErrors verifies the writer's error classification: nil source,
// nil sink, a non-struct row type, and an out-of-range sheet index.
func TestWriteRowsErrors(t *testing.T) {
	rows := []person{{ID: 1, Name: "Ada"}}
	var sink datasource.BytesSink

	if err := editor.WriteRows[person](nil, &sink, rows); !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Errorf("WriteRows(nil source) error = %v, want ErrDataSourceNil", err)
	}
	if err := editor.WriteRows(datasource.BytesSource(newTestWorkbookBytes(t)), nil, rows); !errors.Is(err, toolkiterrors.ErrDataSinkNil) {
		t.Errorf("WriteRows(nil sink) error = %v, want ErrDataSinkNil", err)
	}
	if err := editor.WriteRows(datasource.BytesSource(newTestWorkbookBytes(t)), &sink, []int{1, 2}); !errors.Is(err, toolkiterrors.ErrInvalidValue) {
		t.Errorf("WriteRows with non-struct T error = %v, want ErrInvalidValue", err)
	}
	if err := editor.WriteRows(datasource.BytesSource(newTestWorkbookBytes(t)), &sink, rows, editor.WithSheetIndex(99)); !errors.Is(err, toolkiterrors.ErrInvalidSheetID) {
		t.Errorf("WriteRows WithSheetIndex(99) error = %v, want ErrInvalidSheetID", err)
	}
}

// TestWriteRowsNoHeader verifies that without WithWriteHeader the data starts
// in row 0, so the first cell holds the first field of the first row.
func TestWriteRowsNoHeader(t *testing.T) {
	var sink datasource.BytesSink
	if err := editor.WriteRows(
		datasource.BytesSource(newTestWorkbookBytes(t)),
		&sink,
		[]person{{ID: 5, Name: "Zed"}},
		editor.WithSheetIndex(0),
	); err != nil {
		t.Fatalf("WriteRows: %v", err)
	}
	if err := retryStable(5, func() error {
		v, err := query.ReadCell(datasource.BytesSource(sink.Bytes()), "A1")
		if err != nil {
			return fmt.Errorf("ReadCell A1: %w", err)
		}
		if i, ok := v.Int(); !ok || i != 5 {
			return fmt.Errorf("A1 = %+v, want int 5", v)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

package tests

import (
	"fmt"
	"testing"
	"time"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// The CellValue accessors must refuse a kind mismatch, and an empty cell must
// read as empty. CellValue has no public constructor, so the contract is
// exercised through the exported read path (queryTestWorkbook: A1 is text,
// A3 is an unused cell inside the used range).
func TestCellValueAccessorsRejectMismatchedKind(t *testing.T) {
	src := datasource.BytesSource(queryTestWorkbook(t))
	if err := retryStable(5, func() error {
		text, err := query.ReadCell(src, "A1")
		if err != nil {
			return fmt.Errorf("ReadCell A1: %w", err)
		}
		if text.Kind() != query.KindText {
			return fmt.Errorf("A1 kind = %v, want KindText", text.Kind())
		}
		if _, ok := text.Int(); ok {
			return fmt.Errorf("A1 (text) Int() ok = true, want false")
		}
		if _, ok := text.Float(); ok {
			return fmt.Errorf("A1 (text) Float() ok = true, want false")
		}
		if _, ok := text.Bool(); ok {
			return fmt.Errorf("A1 (text) Bool() ok = true, want false")
		}
		if _, ok := text.Time(); ok {
			return fmt.Errorf("A1 (text) Time() ok = true, want false")
		}
		if text.IsEmpty() {
			return fmt.Errorf("A1 (text) IsEmpty() = true, want false")
		}

		empty, err := query.ReadCell(src, "A3")
		if err != nil {
			return fmt.Errorf("ReadCell A3: %w", err)
		}
		if !empty.IsEmpty() {
			return fmt.Errorf("A3 (empty) IsEmpty() = false, want true")
		}
		if _, ok := empty.String(); ok {
			return fmt.Errorf("A3 (empty) String() ok = true, want false")
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

// String() renders each kind in its canonical text form. The non-date cells are
// deterministic; the date cell is compared against its own Time() rendering so
// the assertion does not depend on engine sub-second precision.
func TestCellValueStringFormats(t *testing.T) {
	src := datasource.BytesSource(queryTestWorkbook(t))
	if err := retryStable(5, func() error {
		cases := []struct {
			ref, want string
		}{
			{"A1", "hello"},
			{"B1", "42"},
			{"C1", "3.14"},
			{"D1", "true"},
		}
		for _, tc := range cases {
			v, err := query.ReadCell(src, tc.ref)
			if err != nil {
				return fmt.Errorf("ReadCell %s: %w", tc.ref, err)
			}
			got, ok := v.String()
			if !ok {
				return fmt.Errorf("%s String() ok = false, want %q", tc.ref, tc.want)
			}
			if got != tc.want {
				return fmt.Errorf("%s String() = %q, want %q", tc.ref, got, tc.want)
			}
		}

		date, err := query.ReadCell(src, "E1")
		if err != nil {
			return fmt.Errorf("ReadCell E1: %w", err)
		}
		ts, ok := date.Time()
		if !ok || ts.Year() != 2024 || ts.Month() != time.January || ts.Day() != 2 {
			return fmt.Errorf("E1 Time = %v, %v; want 2024-01-02, true", ts, ok)
		}
		if got, _ := date.String(); got != ts.Format(time.RFC3339) {
			return fmt.Errorf("E1 String() = %q, want %q", got, ts.Format(time.RFC3339))
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

// Numeric classification: a whole number within int64 range is an integer, a
// whole number beyond it must not wrap an int64 cast and reads as a float.
// Exercised through the engine, so the int64-boundary cells live on their own
// worksheet that no other test's grid depends on.
func TestNumericClassification(t *testing.T) {
	wb, err := asposecells.NewWorkbook()
	if err != nil {
		t.Fatalf("NewWorkbook: %v", err)
	}
	seed, err := wb.Save_SaveFormat(asposecells.SaveFormat_Xlsx)
	if err != nil {
		t.Fatalf("Save_SaveFormat: %v", err)
	}
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(seed),
		editor.InWorksheet(0,
			editor.SetCellValue(0, 0, -7),
			editor.SetCellValue(0, 1, 1e15),
			editor.SetCellValue(0, 2, 1e19),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}
	src := datasource.BytesSource(out)

	if err := retryStable(5, func() error {
		if v, err := query.ReadCell(src, "A1"); err != nil {
			return fmt.Errorf("ReadCell A1: %w", err)
		} else if v.Kind() != query.KindInt {
			return fmt.Errorf("A1 (-7) kind = %v, want KindInt", v.Kind())
		} else if i, _ := v.Int(); i != -7 {
			return fmt.Errorf("A1 Int = %d, want -7", i)
		}

		if v, err := query.ReadCell(src, "B1"); err != nil {
			return fmt.Errorf("ReadCell B1: %w", err)
		} else if v.Kind() != query.KindInt {
			return fmt.Errorf("B1 (1e15) kind = %v, want KindInt", v.Kind())
		} else if i, _ := v.Int(); i != int64(1e15) {
			return fmt.Errorf("B1 Int = %d, want %d", i, int64(1e15))
		}

		// 1e19 is a whole double beyond math.MaxInt64; classification must not
		// wrap it into an integer.
		if v, err := query.ReadCell(src, "C1"); err != nil {
			return fmt.Errorf("ReadCell C1: %w", err)
		} else if v.Kind() != query.KindFloat {
			return fmt.Errorf("C1 (1e19) kind = %v, want KindFloat (beyond int64 range)", v.Kind())
		} else if f, ok := v.Float(); !ok || f <= 0 {
			return fmt.Errorf("C1 Float = %g, %v; want a positive float near 1e19", f, ok)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

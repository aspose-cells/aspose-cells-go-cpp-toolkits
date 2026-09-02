package tests

import (
	"errors"
	"fmt"
	"sync"
	"testing"
	"time"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// queryTestWorkbookBytes caches the built workbook because every call to build
// it costs one engine load, and the evaluation copy caps the total number of
// loads per process (100); rebuilding it per test would exhaust that budget.
var (
	queryTestWorkbookOnce  sync.Once
	queryTestWorkbookBytes []byte
	queryTestWorkbookErr   error
)

// queryTestWorkbook returns a small workbook with known typed content on sheet
// 0 ("Base", renamed at build time), cached so the build happens once:
//
//	       A         B        C         D      E
//	row0   "hello"   42       3.14      true   2024-01-02
//	row1   10        20       =A2+B2
//	row2   (empty)
//	row3   "merged"  (A4:B5 merged region)
//
// Only index-based reads depend on sheet 0 being the first sheet; by-name
// lookups re-run under retryStable because evaluation mode can corrupt a sheet
// name at load time.
func queryTestWorkbook(t *testing.T) []byte {
	t.Helper()
	queryTestWorkbookOnce.Do(func() {
		queryTestWorkbookBytes, queryTestWorkbookErr = buildQueryTestWorkbook()
	})
	if queryTestWorkbookErr != nil {
		t.Fatalf("queryTestWorkbook: %v", queryTestWorkbookErr)
	}
	return queryTestWorkbookBytes
}

func buildQueryTestWorkbook() ([]byte, error) {
	wb, err := asposecells.NewWorkbook()
	if err != nil {
		return nil, err
	}
	wss, err := wb.GetWorksheets()
	if err != nil {
		return nil, err
	}
	first, err := wss.Get_Int(0)
	if err != nil {
		return nil, err
	}
	if err := first.SetName("Base"); err != nil {
		return nil, err
	}
	seed, err := wb.Save_SaveFormat(asposecells.SaveFormat_Xlsx)
	if err != nil {
		return nil, err
	}
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(seed),
		editor.InWorksheet(0,
			editor.SetCellValue(0, 0, "hello"),
			editor.SetCellValue(0, 1, 42),
			editor.SetCellValue(0, 2, 3.14),
			editor.SetCellValue(0, 3, true),
			editor.SetCellValue(0, 4, time.Date(2024, 1, 2, 0, 0, 0, 0, time.UTC)),
			editor.SetCellValue(1, 0, 10),
			editor.SetCellValue(1, 1, 20),
			editor.SetFormula(1, 2, "=A2+B2"),
			editor.Merge(3, 0, 2, 2, false),
			editor.SetCellValue(3, 0, "merged"),
		),
		editor.CalculateAll(),
	)
	if err != nil {
		return nil, err
	}
	return out, nil
}

func TestQueryReadCellTypes(t *testing.T) {
	src := datasource.BytesSource(queryTestWorkbook(t))
	// Read the whole sheet once per attempt instead of issuing one load per
	// cell: every query call re-loads the workbook, and the evaluation copy
	// caps total loads per process.
	if err := retryStable(2, func() error {
		grid, err := query.ReadWorksheet(src)
		if err != nil {
			return fmt.Errorf("ReadWorksheet: %w", err)
		}
		at := func(r, c int) (query.CellValue, error) {
			if r >= len(grid) || c >= len(grid[r]) {
				return query.CellValue{}, fmt.Errorf("grid has no cell (%d, %d)", r, c)
			}
			return grid[r][c], nil
		}

		v, err := at(0, 0)
		if err != nil {
			return err
		}
		if v.Kind() != query.KindText {
			return fmt.Errorf("A1 kind = %v, want KindText", v.Kind())
		}
		if s, _ := v.String(); s != "hello" {
			return fmt.Errorf("A1 = %q, want %q", s, "hello")
		}

		v, err = at(0, 1)
		if err != nil {
			return err
		}
		if v.Kind() != query.KindInt {
			return fmt.Errorf("B1 kind = %v, want KindInt", v.Kind())
		}
		if i, ok := v.Int(); !ok || i != 42 {
			return fmt.Errorf("B1 Int = %d, %v; want 42, true", i, ok)
		}

		v, err = at(0, 2)
		if err != nil {
			return err
		}
		if v.Kind() != query.KindFloat {
			return fmt.Errorf("C1 kind = %v, want KindFloat", v.Kind())
		}
		if f, ok := v.Float(); !ok || f != 3.14 {
			return fmt.Errorf("C1 Float = %f, %v; want 3.14, true", f, ok)
		}

		v, err = at(0, 3)
		if err != nil {
			return err
		}
		if v.Kind() != query.KindBool {
			return fmt.Errorf("D1 kind = %v, want KindBool", v.Kind())
		}
		if b, ok := v.Bool(); !ok || !b {
			return fmt.Errorf("D1 Bool = %v, %v; want true, true", b, ok)
		}

		v, err = at(0, 4)
		if err != nil {
			return err
		}
		if v.Kind() != query.KindDateTime {
			return fmt.Errorf("E1 kind = %v, want KindDateTime", v.Kind())
		}
		if ts, ok := v.Time(); !ok || ts.Year() != 2024 || ts.Month() != time.January || ts.Day() != 2 {
			return fmt.Errorf("E1 Time = %v, %v; want 2024-01-02, true", ts, ok)
		}

		// A formula cell's computed value must be read back after CalculateAll.
		v, err = at(1, 2)
		if err != nil {
			return err
		}
		switch v.Kind() {
		case query.KindInt:
			if i, _ := v.Int(); i != 30 {
				return fmt.Errorf("C2 = %d, want 30", i)
			}
		case query.KindFloat:
			if f, _ := v.Float(); f != 30 {
				return fmt.Errorf("C2 = %f, want 30", f)
			}
		default:
			return fmt.Errorf("C2 kind = %v, want numeric 30", v.Kind())
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestQueryReadRange(t *testing.T) {
	src := datasource.BytesSource(queryTestWorkbook(t))
	if err := retryStable(5, func() error {
		grid, err := query.ReadRange(src, "A1", "B2")
		if err != nil {
			return fmt.Errorf("ReadRange: %w", err)
		}
		if len(grid) != 2 || len(grid[0]) != 2 {
			return fmt.Errorf("ReadRange grid = %d rows x %d cols, want 2 x 2", len(grid), len(grid[0]))
		}
		if s, _ := grid[0][0].String(); grid[0][0].Kind() != query.KindText || s != "hello" {
			return fmt.Errorf("grid[0][0] = %+v, want text hello", grid[0][0])
		}
		if i, ok := grid[0][1].Int(); grid[0][1].Kind() != query.KindInt || !ok || i != 42 {
			return fmt.Errorf("grid[0][1] = %+v, want int 42", grid[0][1])
		}
		if i, ok := grid[1][0].Int(); grid[1][0].Kind() != query.KindInt || !ok || i != 10 {
			return fmt.Errorf("grid[1][0] = %+v, want int 10", grid[1][0])
		}
		if i, ok := grid[1][1].Int(); grid[1][1].Kind() != query.KindInt || !ok || i != 20 {
			return fmt.Errorf("grid[1][1] = %+v, want int 20", grid[1][1])
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestQueryReadWorksheetAndDimensions(t *testing.T) {
	src := datasource.BytesSource(queryTestWorkbook(t))
	if err := retryStable(5, func() error {
		grid, err := query.ReadWorksheet(src)
		if err != nil {
			return fmt.Errorf("ReadWorksheet: %w", err)
		}
		if len(grid) == 0 || len(grid[0]) == 0 {
			return fmt.Errorf("ReadWorksheet grid is empty, want content")
		}
		if s, _ := grid[0][0].String(); s != "hello" {
			return fmt.Errorf("grid[0][0] = %q, want %q", s, "hello")
		}
		rows, cols, err := query.Dimensions(src)
		if err != nil {
			return fmt.Errorf("Dimensions: %w", err)
		}
		if rows == 0 || cols == 0 {
			return fmt.Errorf("Dimensions = (%d, %d), want nonzero", rows, cols)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestQueryReadMergedCells(t *testing.T) {
	src := datasource.BytesSource(queryTestWorkbook(t))
	if err := retryStable(5, func() error {
		areas, err := query.ReadMergedCells(src)
		if err != nil {
			return fmt.Errorf("ReadMergedCells: %w", err)
		}
		want := query.Area{Start: query.CellRef{Row: 3, Col: 0}, End: query.CellRef{Row: 4, Col: 1}}
		for _, a := range areas {
			if a == want {
				// The non-anchor cell of a merged region reads as empty.
				if v, err := query.ReadCell(src, "B5"); err != nil {
					return fmt.Errorf("ReadCell B5: %w", err)
				} else if !v.IsEmpty() {
					return fmt.Errorf("non-anchor merged cell B5 = %+v, want empty", v)
				}
				return nil
			}
		}
		return fmt.Errorf("merged areas = %v, missing %+v", areas, want)
	}); err != nil {
		t.Fatal(err)
	}
}

func TestQuerySheetSelection(t *testing.T) {
	src := datasource.BytesSource(queryTestWorkbook(t))

	// Index-based selection is immune to load-time name corruption, so it must
	// succeed on the first load; still run under retryStable as insurance.
	if err := retryStable(5, func() error {
		v, err := query.ReadCell(src, "A1", query.WithSheetIndex(0))
		if err != nil {
			return fmt.Errorf("WithSheetIndex(0): %w", err)
		}
		if s, _ := v.String(); s != "hello" {
			return fmt.Errorf("WithSheetIndex(0) A1 = %q, want %q", s, "hello")
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}

	// By-name selection needs the "Base" name to survive the load; retry until
	// the engine yields a clean name.
	if err := retryStable(5, func() error {
		v, err := query.ReadCell(src, "A1", query.WithSheet("Base"))
		if err != nil {
			return fmt.Errorf("WithSheet(Base): %w", err)
		}
		if s, _ := v.String(); s != "hello" {
			return fmt.Errorf("WithSheet(Base) A1 = %q, want %q", s, "hello")
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}

	if err := retryStable(5, func() error {
		names, err := query.SheetNames(src)
		if err != nil {
			return fmt.Errorf("SheetNames: %w", err)
		}
		if len(names) == 0 {
			return fmt.Errorf("SheetNames empty, want at least one name")
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestQueryErrors(t *testing.T) {
	src := datasource.BytesSource(queryTestWorkbook(t))

	if _, err := query.ReadCell(nil, "A1"); !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Errorf("ReadCell(nil) error = %v, want ErrDataSourceNil", err)
	}

	if _, err := query.ReadCell(src, "ZZ"); !errors.Is(err, toolkiterrors.ErrInvalidCellRef) {
		t.Errorf("ReadCell(BadRef) error = %v, want ErrInvalidCellRef", err)
	}

	if _, err := query.ReadRange(src, "B2", "A1"); !errors.Is(err, toolkiterrors.ErrInvalidRange) {
		t.Errorf("ReadRange(reversed) error = %v, want ErrInvalidRange", err)
	}

	if _, err := query.ReadCell(src, "A1", query.WithSheet("Nope")); !errors.Is(err, toolkiterrors.ErrWorksheetNotFound) {
		t.Errorf("ReadCell(WithSheet Nope) error = %v, want ErrWorksheetNotFound", err)
	}

	if _, err := query.ReadCell(src, "A1", query.WithSheetIndex(99)); !errors.Is(err, toolkiterrors.ErrInvalidSheetID) {
		t.Errorf("ReadCell(WithSheetIndex 99) error = %v, want ErrInvalidSheetID", err)
	}
}

package tests

import (
	"errors"
	"fmt"
	"strings"
	"sync"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// p2WorkbookBytes caches the workbook shared by the named-range/comment tests,
// built once because every EditSpreadsheet build costs one engine load and the
// evaluation copy caps total loads per process.
var (
	p2WorkbookOnce  sync.Once
	p2WorkbookBytes []byte
	p2WorkbookErr   error
)

func p2Workbook(t *testing.T) []byte {
	t.Helper()
	p2WorkbookOnce.Do(func() {
		p2WorkbookBytes, p2WorkbookErr = buildP2Workbook()
	})
	if p2WorkbookErr != nil {
		t.Fatalf("p2Workbook: %v", p2WorkbookErr)
	}
	return p2WorkbookBytes
}

// buildP2Workbook builds a blank workbook, then on a "Data" sheet added after
// the load defines the named range "Scores" over A1:B2 (values 10, 20, 30, 40)
// and attaches the comment "hello note" to cell A1. The named range's stored
// reference embeds the sheet name, and evaluation mode can rewrite a loaded
// sheet's name to garbage, so the range lives on a sheet whose name is set by
// this build, not read from the file — this keeps the shared build from
// failing on a corrupted default sheet name.
func buildP2Workbook() ([]byte, error) {
	seed, err := newBlankWorkbookBytes()
	if err != nil {
		return nil, err
	}
	return editor.EditSpreadsheet(
		datasource.BytesSource(seed),
		editor.WithAddWorksheet("Data"),
		editor.InWorksheet("Data",
			editor.SetCellValue(0, 0, 10),
			editor.SetCellValue(0, 1, 20),
			editor.SetCellValue(1, 0, 30),
			editor.SetCellValue(1, 1, 40),
			editor.DefineNamedRange("Scores", 0, 0, 1, 1),
			editor.SetCellComment(0, 0, "hello note"),
		),
	)
}

// newBlankWorkbookBytes returns a serialized blank XLSX workbook. NewWorkbook
// creates it in memory (no file opened), so this does not consume the
// evaluation load budget.
func newBlankWorkbookBytes() ([]byte, error) {
	wb, err := asposecells.NewWorkbook()
	if err != nil {
		return nil, err
	}
	return wb.Save_SaveFormat(asposecells.SaveFormat_Xlsx)
}

// TestNamedRangesRoundTrip writes a named range through editor.DefineNamedRange,
// lists it with query.NamedRanges (name, refersTo text, resolved area), and
// reads the referenced cells back with query.ReadNamedRange. The read-back runs
// under retryStable because evaluation mode can corrupt a random cell's value
// at load time.
func TestNamedRangesRoundTrip(t *testing.T) {
	src := datasource.BytesSource(p2Workbook(t))

	if err := retryStable(2, func() error {
		ranges, err := query.NamedRanges(src)
		if err != nil {
			return fmt.Errorf("NamedRanges: %w", err)
		}
		if len(ranges) != 1 {
			return fmt.Errorf("NamedRanges = %d entries (%+v), want 1", len(ranges), ranges)
		}
		nr := ranges[0]
		if nr.Name != "Scores" {
			return fmt.Errorf("name = %q, want %q", nr.Name, "Scores")
		}
		// The refersTo text carries the sheet's name, which evaluation mode may
		// corrupt; assert only the area references it must contain.
		if !strings.Contains(nr.RefersTo, "$A$1") || !strings.Contains(nr.RefersTo, "$B$2") {
			return fmt.Errorf("RefersTo = %q, want it to reference A1:B2", nr.RefersTo)
		}
		want := query.Area{Start: query.CellRef{Row: 0, Col: 0}, End: query.CellRef{Row: 1, Col: 1}}
		if nr.Area != want {
			return fmt.Errorf("area = %+v, want %+v", nr.Area, want)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}

	// ReadNamedRange resolves through the engine's Range object, so it reads the
	// block regardless of sheet selection.
	if err := retryStable(2, func() error {
		grid, err := query.ReadNamedRange(src, "Scores")
		if err != nil {
			return fmt.Errorf("ReadNamedRange: %w", err)
		}
		if len(grid) != 2 || len(grid[0]) != 2 {
			// len(grid[0]) is only safe once len(grid) == 2 is known, so report
			// the row count alone when the grid came back short/empty.
			return fmt.Errorf("ReadNamedRange grid = %d rows, want 2 rows x 2 cols", len(grid))
		}
		want := [2][2]int64{{10, 20}, {30, 40}}
		for r := 0; r < 2; r++ {
			for c := 0; c < 2; c++ {
				v := grid[r][c]
				if iv, ok := v.Int(); v.Kind() != query.KindInt || !ok || iv != want[r][c] {
					return fmt.Errorf("grid[%d][%d] = %+v, want int %d", r, c, v, want[r][c])
				}
			}
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

// TestReadNamedRangeNotFound verifies the error classification for named-range
// reads: a nil source and a name that does not exist.
func TestReadNamedRangeNotFound(t *testing.T) {
	src := datasource.BytesSource(p2Workbook(t))
	if _, err := query.ReadNamedRange(src, "Nope"); !errors.Is(err, toolkiterrors.ErrNameNotFound) {
		t.Errorf("ReadNamedRange(Nope) error = %v, want ErrNameNotFound", err)
	}
	if _, err := query.ReadNamedRange(nil, "Nope"); !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Errorf("ReadNamedRange(nil) error = %v, want ErrDataSourceNil", err)
	}
}

// TestReadCellComment reads a comment written by editor.SetCellComment, verifies
// a comment-free cell reports an empty string, and checks the error
// classification for a nil source and an invalid cell reference. The comment
// lives on the "Data" sheet, which is targeted by name: evaluation mode can
// inject "Evaluation Warning" sheets that shift worksheet indices, so a fixed
// index is unreliable, and can rewrite any loaded sheet's name, so the reads
// are retried until the engine yields a clean "Data" name.
func TestReadCellComment(t *testing.T) {
	src := datasource.BytesSource(p2Workbook(t))
	onData := query.WithSheet("Data")

	if err := retryStable(2, func() error {
		note, err := query.ReadCellComment(src, "A1", onData)
		if err != nil {
			return fmt.Errorf("ReadCellComment(A1): %w", err)
		}
		if note != "hello note" {
			return fmt.Errorf("ReadCellComment(A1) = %q, want %q", note, "hello note")
		}
		empty, err := query.ReadCellComment(src, "B9", onData)
		if err != nil {
			return fmt.Errorf("ReadCellComment(B9): %w", err)
		}
		if empty != "" {
			return fmt.Errorf("ReadCellComment(B9) = %q, want empty", empty)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}

	if _, err := query.ReadCellComment(nil, "A1"); !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Errorf("ReadCellComment(nil) error = %v, want ErrDataSourceNil", err)
	}
	if _, err := query.ReadCellComment(src, "ZZ"); !errors.Is(err, toolkiterrors.ErrInvalidCellRef) {
		t.Errorf("ReadCellComment(ZZ) error = %v, want ErrInvalidCellRef", err)
	}
}

// TestEncrypt verifies editor.Encrypt produces a genuinely encrypted workbook:
// the saved file cannot be loaded without the password, and with the password
// it loads, reports IsEncrypted, and the written value survives.
func TestEncrypt(t *testing.T) {
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(newTestWorkbookBytes(t)),
		editor.InWorksheet(0, editor.SetCellValue(0, 0, int32(42))),
		editor.Encrypt("secret"),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet with Encrypt: %v", err)
	}

	// The saved file is really encrypted: loading it without a password fails.
	if _, err := asposecells.NewWorkbook_Stream(out); err == nil {
		t.Fatal("loading the encrypted workbook without a password should fail")
	}

	// With the password, the workbook loads, reports IsEncrypted, and the value
	// survives the round trip. Retried as insurance against a corrupted cell
	// value at load time.
	if err := retryStable(2, func() error {
		lo, err := asposecells.NewLoadOptions()
		if err != nil {
			return err
		}
		if err := lo.SetPassword("secret"); err != nil {
			return err
		}
		wb, err := asposecells.NewWorkbook_Stream_LoadOptions(out, lo)
		if err != nil {
			return fmt.Errorf("load with password: %w", err)
		}
		settings, err := wb.GetSettings()
		if err != nil {
			return err
		}
		enc, err := settings.IsEncrypted()
		if err != nil {
			return err
		}
		if !enc {
			return fmt.Errorf("IsEncrypted = false, want true")
		}
		wss, err := wb.GetWorksheets()
		if err != nil {
			return err
		}
		ws, err := wss.Get_Int(0)
		if err != nil {
			return err
		}
		cs, err := ws.GetCells()
		if err != nil {
			return err
		}
		cell, err := cs.Get_Int_Int(0, 0)
		if err != nil {
			return err
		}
		iv, err := cell.GetIntValue()
		if err != nil {
			return err
		}
		if iv != 42 {
			return fmt.Errorf("cell value = %d, want 42", iv)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

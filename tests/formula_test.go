package tests

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// TestFormulaWriteCalculateRead writes a formula, recalculates, saves, reloads,
// and confirms both the formula text and its computed value survive the
// round trip. Name corruption is irrelevant (everything is index-based), so the
// retry is insurance only.
func TestFormulaWriteCalculateRead(t *testing.T) {
	if err := retryStable(5, func() error {
		out, err := editor.EditSpreadsheet(
			datasource.BytesSource(newTestWorkbookBytes(t)),
			editor.InWorksheet(0,
				editor.SetCellValue(0, 0, 100),
				editor.SetCellValue(0, 1, 200),
				editor.SetFormula(0, 2, "=A1+B1"),
			),
			editor.CalculateAll(),
		)
		if err != nil {
			return fmt.Errorf("EditSpreadsheet: %w", err)
		}
		wb, err := asposecells.NewWorkbook_Stream(out)
		if err != nil {
			return fmt.Errorf("load: %w", err)
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
		cell, err := cs.Get_Int_Int(0, 2)
		if err != nil {
			return err
		}
		formula, err := cell.GetFormula()
		if err != nil {
			return err
		}
		if formula != "=A1+B1" {
			return fmt.Errorf("formula = %q, want %q", formula, "=A1+B1")
		}
		val, err := cell.GetDoubleValue()
		if err != nil {
			return err
		}
		if val != 300 {
			return fmt.Errorf("computed C1 = %v, want 300", val)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestEditSpreadsheetToSink(t *testing.T) {
	seed := newTestWorkbookBytes(t)
	actions := []editor.WorkbookAction{
		editor.InWorksheet(0, editor.SetCellValue(0, 0, "sink-hello")),
	}

	// The edited output is read back under retryStable: evaluation mode can
	// corrupt a random cell's value at load time (~2% of loads, the same
	// mechanism that corrupts worksheet names), so a single load may spuriously
	// return garbage for A1 even though the bytes are correct.
	t.Run("bytes sink", func(t *testing.T) {
		var sink datasource.BytesSink
		if err := editor.EditSpreadsheetToSink(datasource.BytesSource(seed), &sink, actions...); err != nil {
			t.Fatalf("EditSpreadsheetToSink: %v", err)
		}
		if err := retryStable(5, func() error {
			got, err := readCellString(sink.Bytes(), 0, 0)
			if err != nil {
				return err
			}
			if got != "sink-hello" {
				return fmt.Errorf("bytes sink A1 = %q, want %q", got, "sink-hello")
			}
			return nil
		}); err != nil {
			t.Fatal(err)
		}
	})

	t.Run("file sink", func(t *testing.T) {
		path := filepath.Join(t.TempDir(), "edited.xlsx")
		if err := editor.EditSpreadsheetToSink(datasource.BytesSource(seed), datasource.FilePathSink(path), actions...); err != nil {
			t.Fatalf("EditSpreadsheetToSink to file: %v", err)
		}
		data, err := os.ReadFile(path)
		if err != nil {
			t.Fatalf("ReadFile: %v", err)
		}
		if err := retryStable(5, func() error {
			got, err := readCellString(data, 0, 0)
			if err != nil {
				return err
			}
			if got != "sink-hello" {
				return fmt.Errorf("file sink A1 = %q, want %q", got, "sink-hello")
			}
			return nil
		}); err != nil {
			t.Fatal(err)
		}
	})

	t.Run("EditSpreadsheet equivalent", func(t *testing.T) {
		direct, err := editor.EditSpreadsheet(datasource.BytesSource(seed), actions...)
		if err != nil {
			t.Fatalf("EditSpreadsheet: %v", err)
		}
		if err := retryStable(5, func() error {
			got, err := readCellString(direct, 0, 0)
			if err != nil {
				return err
			}
			if got != "sink-hello" {
				return fmt.Errorf("EditSpreadsheet A1 = %q, want %q", got, "sink-hello")
			}
			return nil
		}); err != nil {
			t.Fatal(err)
		}
	})

	t.Run("nil checks", func(t *testing.T) {
		var sink datasource.BytesSink
		if err := editor.EditSpreadsheetToSink(nil, &sink); !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
			t.Errorf("nil source error = %v, want ErrDataSourceNil", err)
		}
		if err := editor.EditSpreadsheetToSink(datasource.BytesSource(seed), nil); !errors.Is(err, toolkiterrors.ErrDataSinkNil) {
			t.Errorf("nil sink error = %v, want ErrDataSinkNil", err)
		}
	})
}

// readCellString loads a workbook from bytes and returns the string value of
// the cell at the zero-based (row, col) coordinates. It returns an error so
// callers can re-load under retryStable, since evaluation mode can corrupt a
// random cell at load time.
func readCellString(data []byte, row, col int) (string, error) {
	wb, err := asposecells.NewWorkbook_Stream(data)
	if err != nil {
		return "", err
	}
	wss, err := wb.GetWorksheets()
	if err != nil {
		return "", err
	}
	ws, err := wss.Get_Int(0)
	if err != nil {
		return "", err
	}
	cs, err := ws.GetCells()
	if err != nil {
		return "", err
	}
	cell, err := cs.Get_Int_Int(int32(row), int32(col))
	if err != nil {
		return "", err
	}
	s, err := cell.GetStringValue()
	if err != nil {
		return "", err
	}
	return s, nil
}

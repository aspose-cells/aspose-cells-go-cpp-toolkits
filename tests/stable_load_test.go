package tests

import (
	"fmt"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// sheetZeroA1 returns the string value of A1 on the workbook's first sheet.
func sheetZeroA1(wb *asposecells.Workbook) (string, error) {
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
	cell, err := cs.Get_Int_Int(0, 0)
	if err != nil {
		return "", err
	}
	return cell.GetStringValue()
}

// TestLoadStable returns the first workbook whose verification passes, so a
// corrupted load (evaluation-mode name/value corruption) is retried on fresh
// input instead of surfacing garbage.
func TestLoadStable(t *testing.T) {
	src := datasource.BytesSource(queryTestWorkbook(t))
	var verified int
	wb, err := cells.LoadStable(src, 5, func(wb *asposecells.Workbook) error {
		verified++
		got, err := sheetZeroA1(wb)
		if err != nil {
			return err
		}
		if got != "hello" {
			return fmt.Errorf("A1 = %q, want %q", got, "hello")
		}
		return nil
	})
	if err != nil {
		t.Fatalf("LoadStable: %v", err)
	}
	defer wb.Dispose()
	if verified == 0 {
		t.Fatal("verification was never called")
	}
}

// TestLoadStableNoVerify accepts the first load when no verification is given.
func TestLoadStableNoVerify(t *testing.T) {
	src := datasource.BytesSource(queryTestWorkbook(t))
	wb, err := cells.LoadStable(src, 3, nil)
	if err != nil {
		t.Fatalf("LoadStable(nil verify): %v", err)
	}
	defer wb.Dispose()
}

// TestLoadStableExhausted reports an error and stops after attempts retries when
// every load fails verification.
func TestLoadStableExhausted(t *testing.T) {
	src := datasource.BytesSource(queryTestWorkbook(t))
	const attempts = 3
	var calls int
	_, err := cells.LoadStable(src, attempts, func(*asposecells.Workbook) error {
		calls++
		return fmt.Errorf("verification always fails")
	})
	if err == nil {
		t.Fatal("LoadStable with always-failing verification: want error, got nil")
	}
	if calls != attempts {
		t.Errorf("verification called %d times, want %d", calls, attempts)
	}
}

package tests

import (
	"errors"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/transfer"
)

// Referencing a worksheet by a name that does not exist must return an error,
// not crash. The binding's Get_String returns err=nil with a dangling native
// handle for missing names, so every name-based lookup iterates the collection
// instead. These tests lock in that behaviour at the public entry points that
// used to segfault the whole process.
func TestMissingWorksheetReturnsErrorNotCrash(t *testing.T) {
	source := datasource.BytesSource(newTestWorkbookBytes(t))

	t.Run("editor rename", func(t *testing.T) {
		_, err := editor.EditSpreadsheet(source,
			editor.WithRenameWorksheet("NoSuchSheet", "Renamed"),
		)
		if err == nil {
			t.Fatal("WithRenameWorksheet with a missing sheet: want error, got nil")
		}
		if !errors.Is(err, toolkiterrors.ErrWorksheetNotFound) {
			t.Errorf("error = %v, want wrapping of ErrWorksheetNotFound", err)
		}
	})

	t.Run("editor in-worksheet", func(t *testing.T) {
		_, err := editor.EditSpreadsheet(source,
			editor.InWorksheet("NoSuchSheet", editor.SetCellValue(0, 0, "x")),
		)
		if err == nil {
			t.Fatal("InWorksheet with a missing sheet: want error, got nil")
		}
		if !errors.Is(err, toolkiterrors.ErrWorksheetNotFound) {
			t.Errorf("error = %v, want wrapping of ErrWorksheetNotFound", err)
		}
	})

	t.Run("transfer json", func(t *testing.T) {
		_, err := transfer.ExportWorksheetToJson(source, "NoSuchSheet")
		if err == nil {
			t.Fatal("ExportWorksheetToJson with a missing sheet: want error, got nil")
		}
		if !errors.Is(err, toolkiterrors.ErrWorksheetNotFound) {
			t.Errorf("error = %v, want wrapping of ErrWorksheetNotFound", err)
		}
	})

	t.Run("transfer range json", func(t *testing.T) {
		_, err := transfer.ExportRangeToJson(source, "NoSuchSheet", "A1", "B2")
		if err == nil {
			t.Fatal("ExportRangeToJson with a missing sheet: want error, got nil")
		}
		if !errors.Is(err, toolkiterrors.ErrWorksheetNotFound) {
			t.Errorf("error = %v, want wrapping of ErrWorksheetNotFound", err)
		}
	})
}

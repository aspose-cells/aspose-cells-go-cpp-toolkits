package tests

import (
	"errors"
	"os"
	"path/filepath"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/transfer"
)

// TestSentinelErrors verifies the toolkit's sentinel errors are distinguishable
// with errors.Is, so callers never have to match error message strings.
func TestSentinelErrors(t *testing.T) {
	t.Run("unsupported format", func(t *testing.T) {
		// Extension check happens before the input file is opened, so the
		// input path need not exist.
		err := converter.ConvertSpreadsheetToFile("input.xlsx", "out.zzz9")
		if !errors.Is(err, toolkiterrors.ErrUnsupportedFormat) {
			t.Fatalf("ConvertSpreadsheetToFile error = %v, want ErrUnsupportedFormat", err)
		}
	})

	t.Run("nil save option", func(t *testing.T) {
		_, err := converter.ConvertSpreadsheet(datasource.BytesSource([]byte{}), nil)
		if !errors.Is(err, toolkiterrors.ErrSaveOptionNil) {
			t.Fatalf("ConvertSpreadsheet error = %v, want ErrSaveOptionNil", err)
		}
	})

	t.Run("invalid value", func(t *testing.T) {
		_, err := editor.EditSpreadsheet(
			datasource.BytesSource(newTestWorkbookBytes(t)),
			editor.InWorksheet(0, editor.SetCellValue(0, 0, complex(1, 2))),
		)
		if !errors.Is(err, toolkiterrors.ErrInvalidValue) {
			t.Fatalf("EditSpreadsheet error = %v, want ErrInvalidValue", err)
		}
	})

	t.Run("input is folder", func(t *testing.T) {
		dir := t.TempDir()
		err := transfer.ExportSpreadsheetToXmlFile(dir, "map", filepath.Join(dir, "out.xml"))
		if !errors.Is(err, toolkiterrors.ErrInputIsFolder) {
			t.Fatalf("ExportSpreadsheetToXmlFile error = %v, want ErrInputIsFolder", err)
		}
	})

	t.Run("missing input is not a format error", func(t *testing.T) {
		// A missing input file must surface the os error (errors.Is with
		// fs.ErrNotExist) and never be misclassified as ErrUnsupportedFormat.
		err := converter.ConvertSpreadsheetToFile(filepath.Join(t.TempDir(), "missing.xlsx"), "out.pdf")
		if errors.Is(err, toolkiterrors.ErrUnsupportedFormat) {
			t.Fatalf("missing input misclassified as ErrUnsupportedFormat: %v", err)
		}
		if err == nil || !os.IsNotExist(err) {
			t.Fatalf("ConvertSpreadsheetToFile missing file error = %v, want fs.ErrNotExist", err)
		}
	})
}

package tests

import (
	"fmt"
	"strings"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	_ "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/register"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/image"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// TestFormatDiscriminators verifies each registered format factory reports the
// exact format string it was registered under. This guards against a factory
// registering several names but always building the default option — the ooxml
// xlsm/xltx/xltm bug, where GetFormat() hardcoded "xlsx" and Apply always used
// the default XLSX constructor, and the txt tsv gap, where "tsv" was never
// registered at all.
func TestFormatDiscriminators(t *testing.T) {
	cases := []struct{ ext, want string }{
		{"xlsm", "xlsm"},
		{"xltx", "xltx"},
		{"xltm", "xltm"},
		{"tsv", "tsv"},
		{"jpeg", "jpeg"},
		{"gif", "gif"},
		{"emf", "emf"},
		{"png", "png"},
		{"jpg", "jpg"},
	}
	for _, tc := range cases {
		t.Run(tc.ext, func(t *testing.T) {
			opt := formats.Get(tc.ext)
			if opt == nil {
				t.Fatalf("format %q is not registered", tc.ext)
			}
			if got := opt.GetFormat(); got != tc.want {
				t.Errorf("Get(%q).GetFormat() = %q, want %q", tc.ext, got, tc.want)
			}
		})
	}
}

// TestImageDefaultFormatIsPng verifies image.New() without options reports
// "png", so converter.Convert(..., image.New(), ...) writes a PNG rather than
// a file with no recognized extension.
func TestImageDefaultFormatIsPng(t *testing.T) {
	if got := image.New().GetFormat(); got != "png" {
		t.Errorf("image.New().GetFormat() = %q, want %q", got, "png")
	}
}

// TestXlsmRoundTripProducesXlsm is the behavioral half of the ooxml
// discriminator regression: formats.Get("xlsm") must write XLSM (macro-enabled
// workbook) bytes, not plain XLSX. Before the fix the option reported "xlsm"
// but Apply always used the default OOXML save format.
func TestXlsmRoundTripProducesXlsm(t *testing.T) {
	if err := retryStable(5, func() error {
		src := datasource.BytesSource(newNamedWorkbookBytes(t, "M"))
		var out datasource.BytesSink
		if err := converter.Convert(src, formats.Get("xlsm"), &out); err != nil {
			return fmt.Errorf("Convert to xlsm: %w", err)
		}
		wb, err := asposecells.NewWorkbook_Stream(out.Bytes())
		if err != nil {
			return fmt.Errorf("load converted bytes: %w", err)
		}
		got, err := wb.GetFileFormat()
		if err != nil {
			return fmt.Errorf("GetFileFormat: %w", err)
		}
		if got != asposecells.FileFormatType_Xlsm {
			return fmt.Errorf("Convert(Get(xlsm)) file format = %v, want %v", got, asposecells.FileFormatType_Xlsm)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

// newWorkbookWithValue builds a workbook whose first sheet holds value in cell
// A1 and label in B1, so a delimited-text export is guaranteed non-empty.
func newWorkbookWithValue(t *testing.T, name, value, label string) []byte {
	t.Helper()
	wb, err := asposecells.NewWorkbook()
	if err != nil {
		t.Fatalf("NewWorkbook: %v", err)
	}
	wss, err := wb.GetWorksheets()
	if err != nil {
		t.Fatalf("GetWorksheets: %v", err)
	}
	ws, err := wss.Get_Int(0)
	if err != nil {
		t.Fatalf("Get_Int(0): %v", err)
	}
	if err := ws.SetName(name); err != nil {
		t.Fatalf("SetName: %v", err)
	}
	cells, err := ws.GetCells()
	if err != nil {
		t.Fatalf("GetCells: %v", err)
	}
	values := []string{value, label}
	for i, v := range values {
		cell, err := cells.Get_Int_Int(0, int32(i))
		if err != nil {
			t.Fatalf("Get_Int_Int(0,%d): %v", i, err)
		}
		obj, err := asposecells.NewObject_String(v)
		if err != nil {
			t.Fatalf("NewObject_String(%q): %v", v, err)
		}
		if err := cell.PutValue_Object(obj); err != nil {
			t.Fatalf("PutValue_Object(%q): %v", v, err)
		}
	}
	data, err := wb.Save_SaveFormat(asposecells.SaveFormat_Xlsx)
	if err != nil {
		t.Fatalf("Save_SaveFormat: %v", err)
	}
	return data
}

// TestTsvRoundTripIsTabSeparated verifies formats.Get("tsv") actually emits
// tab-separated text. The check reads the cell values back out of the saved
// bytes, so a corrupted worksheet name in evaluation mode cannot fail it.
func TestTsvRoundTripIsTabSeparated(t *testing.T) {
	src := datasource.BytesSource(newWorkbookWithValue(t, "Data", "alpha", "beta"))
	var out datasource.BytesSink
	if err := converter.Convert(src, formats.Get("tsv"), &out); err != nil {
		t.Fatalf("Convert to tsv: %v", err)
	}
	got := string(out.Bytes())
	if !strings.Contains(got, "alpha") || !strings.Contains(got, "beta") || !strings.Contains(got, "\t") {
		t.Errorf("Convert(Get(tsv)) output = %q, want values alpha/beta separated by a tab", got)
	}
}

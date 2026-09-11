package tests

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/manipulator"
	_ "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/register"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/transfer"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// newNamedWorkbookBytes builds a workbook holding the default first sheet plus
// one sheet per given name. In evaluation mode the engine can corrupt the
// default first sheet's name, so the first sheet is renamed by index to a fixed
// safe name ("Base") at build time. Note the corruption is not limited to the
// default sheet and not captured in the bytes: a probe showed the same bytes
// loading clean then, on a second load, with sheet[0] renamed to garbage
// ("\x00@\x12\x00"). It is a load-time roll of the dice (~2% per load, any
// sheet). Callers that must depend on surviving names run inside retryStable,
// which re-rolls the whole operation on fresh input.
func newNamedWorkbookBytes(t *testing.T, names ...string) []byte {
	t.Helper()
	wb, err := asposecells.NewWorkbook()
	if err != nil {
		t.Fatalf("NewWorkbook: %v", err)
	}
	wss, err := wb.GetWorksheets()
	if err != nil {
		t.Fatalf("GetWorksheets: %v", err)
	}
	first, err := wss.Get_Int(0)
	if err != nil {
		t.Fatalf("Get_Int(0): %v", err)
	}
	if err := first.SetName("Base"); err != nil {
		t.Fatalf("rename first sheet: %v", err)
	}
	for _, n := range names {
		if _, err := wss.Add_String(n); err != nil {
			t.Fatalf("Add_String(%q): %v", n, err)
		}
	}
	data, err := wb.Save_SaveFormat(asposecells.SaveFormat_Xlsx)
	if err != nil {
		t.Fatalf("Save_SaveFormat: %v", err)
	}
	return data
}

// retryStable runs fn until it returns nil, up to max attempts, and returns the
// last error if every attempt fails. In evaluation mode the engine corrupts a
// random worksheet name at load time with small probability (~2% per load, any
// sheet), which can spuriously fail operations whose success depends on
// surviving names (by-name lookups, per-sheet split outputs). Re-running with
// freshly generated input re-rolls the engine, while a genuine deterministic
// defect (for example Split always dropping a sheet) fails every attempt and
// still surfaces. The caller must call t.Fatal on the returned error — it must
// not fail from within fn, because a retry loop needs the engine roll and the
// failure to be retryable.
func retryStable(max int, fn func() error) error {
	var last error
	for attempt := 1; attempt <= max; attempt++ {
		if err := fn(); err != nil {
			last = fmt.Errorf("attempt %d: %w", attempt, err)
			continue
		}
		return nil
	}
	return last
}

func TestMergeEmptySourcesErrors(t *testing.T) {
	opt := formats.Get("xlsx")
	var out datasource.BytesSink

	t.Run("nil sources", func(t *testing.T) {
		err := manipulator.Merge(nil, opt, &out)
		if !errors.Is(err, toolkiterrors.ErrNoSources) {
			t.Fatalf("Merge(nil) error = %v, want ErrNoSources", err)
		}
	})

	t.Run("empty sources", func(t *testing.T) {
		err := manipulator.Merge([]datasource.DataSource{}, opt, &out)
		if !errors.Is(err, toolkiterrors.ErrNoSources) {
			t.Fatalf("Merge([]) error = %v, want ErrNoSources", err)
		}
	})

	t.Run("nil element carries index", func(t *testing.T) {
		err := manipulator.Merge([]datasource.DataSource{nil}, opt, &out)
		if !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
			t.Fatalf("Merge([nil]) error = %v, want ErrDataSourceNil", err)
		}
		if !strings.Contains(err.Error(), "source 0") {
			t.Errorf("Merge([nil]) error %q should mention source index 0", err)
		}
	})
}

func TestMergeCombinesWorkbooks(t *testing.T) {
	if err := retryStable(5, func() error {
		alpha := datasource.BytesSource(newNamedWorkbookBytes(t, "Alpha"))
		beta := datasource.BytesSource(newNamedWorkbookBytes(t, "Beta"))

		var out datasource.BytesSink
		if err := manipulator.Merge([]datasource.DataSource{alpha, beta}, formats.Get("xlsx"), &out); err != nil {
			return fmt.Errorf("Merge: %w", err)
		}
		merged, err := asposecells.NewWorkbook_Stream(out.Bytes())
		if err != nil {
			return fmt.Errorf("load merged workbook: %w", err)
		}
		wss, err := merged.GetWorksheets()
		if err != nil {
			return fmt.Errorf("GetWorksheets: %w", err)
		}
		count, err := wss.GetCount()
		if err != nil {
			return fmt.Errorf("GetCount: %w", err)
		}
		if count < 2 {
			return fmt.Errorf("merged workbook has %d worksheets, want >= 2 (one per source)", count)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestSplitProducesOneOutputPerSheet(t *testing.T) {
	// Evaluation mode can append an extra "Evaluation Warning" sheet, so exact
	// total counts are unreliable; instead assert that each explicitly named
	// worksheet gets its own non-empty output file / archive entry. A worksheet
	// name can also be corrupted at load time, so each split below is re-run on
	// a fresh source until the engine yields clean names (retryStable).
	opt := formats.Get("xlsx")
	want := []string{"A.xlsx", "B.xlsx", "C.xlsx"}

	t.Run("folder sink", func(t *testing.T) {
		if err := retryStable(5, func() error {
			src := datasource.BytesSource(newNamedWorkbookBytes(t, "A", "B", "C"))
			dir := t.TempDir()
			if err := manipulator.Split(src, opt, datasource.FolderSink(dir)); err != nil {
				return fmt.Errorf("Split to folder: %w", err)
			}
			entries, err := os.ReadDir(dir)
			if err != nil {
				return fmt.Errorf("ReadDir: %w", err)
			}
			byName := map[string]bool{}
			for _, e := range entries {
				byName[e.Name()] = true
				info, err := e.Info()
				if err != nil {
					return fmt.Errorf("entry info: %w", err)
				}
				if info.Size() == 0 {
					return fmt.Errorf("split file %q is empty", e.Name())
				}
			}
			for _, name := range want {
				if !byName[name] {
					return fmt.Errorf("split folder missing %q; got %v", name, entries)
				}
			}
			return nil
		}); err != nil {
			t.Fatal(err)
		}
	})

	t.Run("zip sink", func(t *testing.T) {
		if err := retryStable(5, func() error {
			src := datasource.BytesSource(newNamedWorkbookBytes(t, "A", "B", "C"))
			buf := new(bytes.Buffer)
			zw := zip.NewWriter(buf)
			if err := manipulator.Split(src, opt, datasource.NewZipSink(zw)); err != nil {
				return fmt.Errorf("Split to zip: %w", err)
			}
			if err := zw.Close(); err != nil {
				return fmt.Errorf("zip close: %w", err)
			}
			zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
			if err != nil {
				return fmt.Errorf("zip reader: %w", err)
			}
			names := map[string]bool{}
			for _, f := range zr.File {
				names[f.Name] = true
			}
			for _, name := range want {
				if !names[name] {
					return fmt.Errorf("zip missing entry %q; got %v", name, names)
				}
			}
			return nil
		}); err != nil {
			t.Fatal(err)
		}
	})
}

func TestExportEmptySheetWritesEmptyJson(t *testing.T) {
	if err := retryStable(5, func() error {
		src := datasource.BytesSource(newNamedWorkbookBytes(t, "EmptyData"))
		var out datasource.BytesSink
		if err := transfer.ExportWorksheetToJson(src, &out, transfer.WithSheet("EmptyData")); err != nil {
			return fmt.Errorf("ExportWorksheetToJson on empty sheet: %w", err)
		}
		if got := string(out.Bytes()); got != "[]" {
			return fmt.Errorf("empty sheet JSON = %q, want %q", got, "[]")
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestImportCSVLandsInCells(t *testing.T) {
	if err := retryStable(5, func() error {
		seed := datasource.BytesSource(newNamedWorkbookBytes(t, "Imported"))
		csv := datasource.BytesSource([]byte("id,name\n1,Alpha\n2,Beta\n"))

		var imported datasource.BytesSink
		if err := transfer.ImportCSV(seed, csv, &imported,
			transfer.WithSheet("Imported"), transfer.WithBeginCell(0, 0),
			transfer.WithConvertNumeric(true), transfer.WithSeparator(","),
		); err != nil {
			return fmt.Errorf("ImportCSV: %w", err)
		}

		// Round-trip: export the imported sheet as JSON and confirm the CSV values
		// actually landed in cells.
		var jsonOut datasource.BytesSink
		if err := transfer.ExportWorksheetToJson(datasource.BytesSource(imported.Bytes()), &jsonOut, transfer.WithSheet("Imported")); err != nil {
			return fmt.Errorf("export imported sheet: %w", err)
		}
		got := string(jsonOut.Bytes())
		for _, want := range []string{"Alpha", "Beta"} {
			if !strings.Contains(got, want) {
				return fmt.Errorf("imported sheet JSON missing %q:\n%s", want, got)
			}
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestFilePathSinkCreatesParentDir(t *testing.T) {
	path := filepath.Join(t.TempDir(), "nested", "dir", "out.txt")
	if err := datasource.FilePathSink(path).Write("", []byte("content")); err != nil {
		t.Fatalf("FilePathSink.Write: %v", err)
	}
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("ReadFile: %v", err)
	}
	if string(data) != "content" {
		t.Errorf("file content = %q, want %q", data, "content")
	}
}

package cells

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// SaveWorkbookToSink serializes workbook in its own loaded format (see
// WorkbookToByteData) and writes the result to sink.
//
// A datasource.FilePathSink is written straight to disk by the engine: the
// document is saved into a temporary file in the destination directory and
// renamed into place, matching FilePathSink.Write's atomic-replacement
// semantics. The serialized document therefore never has to be held in Go
// memory, which halves the peak for a large edit-to-file. Every other sink
// gets the bytes through Write, exactly as before.
func SaveWorkbookToSink(workbook *asposecells.Workbook, sink datasource.DataSink) error {
	fileFormat, err := workbook.GetFileFormat()
	if err != nil {
		return err
	}
	saveFormat, err := engine.FileFormatToSaveFormat(fileFormat)
	if err != nil {
		return err
	}
	if path, ok := sink.(datasource.FilePathSink); ok {
		return saveWorkbookToFile(workbook, string(path), saveFormat)
	}
	data, err := workbook.Save_SaveFormat(saveFormat)
	if err != nil {
		return err
	}
	return sink.Write("", data)
}

// saveWorkbookToFile saves workbook to path through a temporary file in the
// same directory, then renames it over the destination. A crash or a full disk
// mid-save leaves the previous file intact, the same guarantee
// datasource.FilePathSink.Write gives.
func saveWorkbookToFile(workbook *asposecells.Workbook, path string, saveFormat asposecells.SaveFormat) error {
	dir := filepath.Dir(path)
	if dir != "" && dir != "." {
		if err := os.MkdirAll(dir, 0o755); err != nil {
			return fmt.Errorf("create output directory %q: %w", dir, err)
		}
	}
	tmp, err := os.CreateTemp(dir, "."+filepath.Base(path)+".tmp*")
	if err != nil {
		return fmt.Errorf("write %q: %w", path, err)
	}
	tmpName := tmp.Name()
	// The engine opens the temp path itself, so the Go handle is only a name
	// reservation; holding it open would share the file with the native write.
	if err := tmp.Close(); err != nil {
		os.Remove(tmpName)
		return fmt.Errorf("write %q: %w", path, err)
	}
	if err := workbook.Save_String_SaveFormat(tmpName, saveFormat); err != nil {
		os.Remove(tmpName)
		return err
	}
	if err := os.Chmod(tmpName, 0o644); err != nil {
		os.Remove(tmpName)
		return fmt.Errorf("write %q: %w", path, err)
	}
	if err := os.Rename(tmpName, path); err != nil {
		os.Remove(tmpName)
		return fmt.Errorf("write %q: %w", path, err)
	}
	return nil
}

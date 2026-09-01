// Package cells provides shared helpers over the Aspose.Cells engine used by
// the toolkit's public packages.
//
// It centralizes reading a datasource.DataSource into bytes, building a
// Workbook from it, serializing a Workbook back to bytes, and resolving cells
// and cell areas, so the converter, editor, manipulator, and transfer packages
// do not each re-implement the same engine plumbing.
package cells

import (
	"fmt"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"io"
)

// ReadSource reads all bytes from a datasource.DataSource. It opens the source,
// drains it fully into memory, and closes it. This is the single place where a
// DataSource is turned into raw bytes, so every package reads sources the same way.
func ReadSource(source datasource.DataSource) ([]byte, error) {
	reader, errOpen := source.Open()
	if errOpen != nil {
		return nil, errOpen
	}
	defer reader.Close()
	data, errRead := io.ReadAll(reader)
	if errRead != nil {
		return nil, errRead
	}
	return data, nil
}

func GetWorkbookWithDataSource(source datasource.DataSource) (*asposecells.Workbook, error) {
	data, err := ReadSource(source)
	if err != nil {
		return nil, err
	}
	return asposecells.NewWorkbook_Stream(data)
}

func WorkbookToByteData(workbook *asposecells.Workbook) ([]byte, error) {
	fileFormat, err := workbook.GetFileFormat()
	if err != nil {
		return nil, err
	}
	saveFormat := formats.FileFormatToSaveFormat(fileFormat)
	return workbook.Save_SaveFormat(saveFormat)
}

// LoadStable loads a workbook from source, retrying on fresh loads when the
// engine corrupts state at load time. In evaluation mode NewWorkbook_Stream
// occasionally returns a workbook whose in-memory state (a worksheet name or a
// cell value) is garbage — ~2% of loads, non-deterministic, and not present in
// the bytes, so the same source can load clean once and corrupt later.
//
// verify is called on each loaded workbook; when it reports a non-nil error the
// load is retried from a fresh source read, up to attempts times. A nil verify
// accepts the first load. Every rejected workbook is disposed so the corrupted
// engine state does not leak a native handle.
//
// The source is re-opened per attempt, so all datasource types work: a file is
// re-read, a BytesSource hands out a fresh reader, and a ReaderSource serves
// its buffered bytes.
func LoadStable(source datasource.DataSource, attempts int, verify func(*asposecells.Workbook) error) (*asposecells.Workbook, error) {
	if attempts < 1 {
		attempts = 1
	}
	var lastErr error
	for i := 0; i < attempts; i++ {
		wb, err := GetWorkbookWithDataSource(source)
		if err != nil {
			// A failed load (e.g. invalid bytes) is deterministic; retrying
			// cannot help, so report it immediately.
			return nil, err
		}
		if verify == nil {
			return wb, nil
		}
		if err := verify(wb); err != nil {
			lastErr = err
			_ = wb.Dispose()
			continue
		}
		return wb, nil
	}
	return nil, fmt.Errorf("workbook failed verification after %d loads: %w", attempts, lastErr)
}

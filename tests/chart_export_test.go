package tests

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"os"
	"sync"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
)

var (
	chartExportFixtureOnce sync.Once
	chartExportFixtureData []byte
	chartExportFixtureErr  error
)

// getChartExportFixture returns a workbook with one column chart over a small
// table. It is built once per process because every EditSpreadsheet costs one
// of the evaluation copy's per-process loads.
func getChartExportFixture(t *testing.T) []byte {
	t.Helper()
	chartExportFixtureOnce.Do(func() {
		seed, err := datasource.NewEmptyWorkbook()
		if err != nil {
			chartExportFixtureErr = err
			return
		}
		out, err := editor.EditSpreadsheet(
			seed,
			editor.InWorksheet(0,
				editor.SetCellValue(0, 0, "Category"),
				editor.SetCellValue(0, 1, "Value"),
				editor.SetCellValue(1, 0, "A"),
				editor.SetCellValue(1, 1, 10),
				editor.SetCellValue(2, 0, "B"),
				editor.SetCellValue(2, 1, 20),
				editor.SetCellValue(3, 0, "C"),
				editor.SetCellValue(3, 1, 30),
				editor.AddChart(editor.ChartTypeColumn, "A1:B4", false, 0, 3, 15, 10,
					editor.WithChartTitle("Test Chart"),
					editor.WithChartStyle(7),
				),
			),
		)
		if err != nil {
			chartExportFixtureErr = err
			return
		}
		chartExportFixtureData = out
	})
	if chartExportFixtureErr != nil {
		t.Fatalf("create chart fixture: %v", chartExportFixtureErr)
	}
	return chartExportFixtureData
}

// pngDimensions reads a PNG's pixel size from its IHDR chunk.
func pngDimensions(t *testing.T, data []byte) (width, height int) {
	t.Helper()
	if len(data) < 24 || string(data[12:16]) != "IHDR" {
		t.Fatalf("exported data is not a PNG (len=%d)", len(data))
	}
	return int(binary.BigEndian.Uint32(data[16:20])), int(binary.BigEndian.Uint32(data[20:24]))
}

func isPNG(data []byte) bool {
	return len(data) >= 8 && bytes.Equal(data[:8], []byte{0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A})
}

func isJPEG(data []byte) bool {
	return len(data) >= 2 && data[0] == 0xFF && data[1] == 0xD8
}

func isPDF(data []byte) bool {
	return bytes.HasPrefix(data, []byte("%PDF-"))
}

// --- format tests ---------------------------------------------------------

func TestExportChartToBytes_PNG(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	data, err := converter.ExportChartToBytes(src, 0, 0, &converter.ChartExportOptions{
		Format: converter.ChartExportFormatPNG,
	})
	if err != nil {
		t.Fatalf("ExportChartToBytes PNG: %v", err)
	}
	if !isPNG(data) {
		t.Fatal("exported data is not PNG")
	}
	width, height := pngDimensions(t, data)
	if width <= 0 || height <= 0 {
		t.Fatalf("PNG reports implausible size %dx%d", width, height)
	}
}

func TestExportChartToBytes_JPEG(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	data, err := converter.ExportChartToBytes(src, 0, 0, &converter.ChartExportOptions{
		Format:  converter.ChartExportFormatJPEG,
		Quality: 85,
	})
	if err != nil {
		t.Fatalf("ExportChartToBytes JPEG: %v", err)
	}
	if !isJPEG(data) {
		t.Fatal("exported data is not JPEG")
	}
}

func TestExportChartToBytes_SVG(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	data, err := converter.ExportChartToBytes(src, 0, 0, &converter.ChartExportOptions{
		Format: converter.ChartExportFormatSVG,
	})
	if err != nil {
		t.Fatalf("ExportChartToBytes SVG: %v", err)
	}
	if len(data) == 0 {
		t.Fatal("exported SVG is empty")
	}
	if !bytes.Contains(data, []byte("<svg")) {
		t.Fatal("exported data contains no <svg> element")
	}
}

// TestExportChartToBytes_PDF checks the PDF path produces a real PDF document,
// not the EMF bytes that a chart-to-image call yields.
func TestExportChartToBytes_PDF(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	data, err := converter.ExportChartToBytes(src, 0, 0, &converter.ChartExportOptions{
		Format: converter.ChartExportFormatPDF,
	})
	if err != nil {
		t.Fatalf("ExportChartToBytes PDF: %v", err)
	}
	if !isPDF(data) {
		head := data
		if len(head) > 16 {
			head = head[:16]
		}
		t.Fatalf("exported data is not a PDF, header = %q", head)
	}
	if !bytes.Contains(data, []byte("%%EOF")) {
		t.Error("PDF is missing its EOF trailer, so it is likely truncated")
	}
}

// --- sizing ---------------------------------------------------------------

// TestExportChartToBytes_ExactDimensions is the test that makes the size
// option meaningful: a requested size must be the delivered size, exactly.
func TestExportChartToBytes_ExactDimensions(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	for _, size := range []struct{ width, height int }{
		{320, 240},
		{800, 600},
		{1200, 400},
	} {
		t.Run(fmt.Sprintf("%dx%d", size.width, size.height), func(t *testing.T) {
			data, err := converter.ExportChartToBytes(src, 0, 0, &converter.ChartExportOptions{
				Format: converter.ChartExportFormatPNG,
				Width:  size.width,
				Height: size.height,
			})
			if err != nil {
				t.Fatalf("ExportChartToBytes %dx%d: %v", size.width, size.height, err)
			}
			gotWidth, gotHeight := pngDimensions(t, data)
			if gotWidth != size.width || gotHeight != size.height {
				t.Errorf("exported size = %dx%d, want %dx%d",
					gotWidth, gotHeight, size.width, size.height)
			}
		})
	}
}

// TestExportChartToBytes_DefaultSize leaves both dimensions at zero, which
// means "use the chart's own size" and must not be an error.
func TestExportChartToBytes_DefaultSize(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	data, err := converter.ExportChartToBytes(src, 0, 0, nil)
	if err != nil {
		t.Fatalf("ExportChartToBytes with no options: %v", err)
	}
	if !isPNG(data) {
		t.Fatal("default format should be PNG")
	}
	if width, height := pngDimensions(t, data); width <= 0 || height <= 0 {
		t.Fatalf("default-size PNG reports %dx%d", width, height)
	}
}

// TestExportChartToBytes_PartialSize rejects a half-specified size rather than
// guessing the missing dimension.
func TestExportChartToBytes_PartialSize(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	for _, tc := range []struct {
		name          string
		width, height int
	}{
		{"width only", 800, 0},
		{"height only", 0, 600},
	} {
		t.Run(tc.name, func(t *testing.T) {
			_, err := converter.ExportChartToBytes(src, 0, 0, &converter.ChartExportOptions{
				Format: converter.ChartExportFormatPNG,
				Width:  tc.width,
				Height: tc.height,
			})
			if !errors.Is(err, toolkiterrors.ErrInvalidChartSize) {
				t.Fatalf("error = %v, want ErrInvalidChartSize", err)
			}
		})
	}
}

func TestExportChartToBytes_NegativeSize(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	_, err := converter.ExportChartToBytes(src, 0, 0, &converter.ChartExportOptions{
		Format: converter.ChartExportFormatPNG,
		Width:  -10,
		Height: 100,
	})
	if !errors.Is(err, toolkiterrors.ErrInvalidChartSize) {
		t.Fatalf("error = %v, want ErrInvalidChartSize", err)
	}
}

// --- quality --------------------------------------------------------------

// TestExportChartToBytes_JPEGQuality checks the quality setting actually
// reaches the encoder: a low quality must produce a smaller file than a high
// one, and both must still be valid JPEG.
func TestExportChartToBytes_JPEGQuality(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	export := func(quality int) []byte {
		t.Helper()
		data, err := converter.ExportChartToBytes(src, 0, 0, &converter.ChartExportOptions{
			Format:  converter.ChartExportFormatJPEG,
			Quality: quality,
		})
		if err != nil {
			t.Fatalf("ExportChartToBytes JPEG quality %d: %v", quality, err)
		}
		if !isJPEG(data) {
			t.Fatalf("quality %d output is not JPEG", quality)
		}
		return data
	}

	low := export(10)
	high := export(95)

	if bytes.Equal(low, high) {
		t.Error("quality 10 and quality 95 produced identical bytes, so quality is not applied")
	}
	if len(high) <= len(low) {
		t.Errorf("quality 95 produced %d bytes, quality 10 produced %d; high quality should be larger",
			len(high), len(low))
	}
}

func TestExportChartToBytes_InvalidQuality(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	for _, quality := range []int{-1, 101, 1000} {
		t.Run(fmt.Sprintf("quality_%d", quality), func(t *testing.T) {
			_, err := converter.ExportChartToBytes(src, 0, 0, &converter.ChartExportOptions{
				Format:  converter.ChartExportFormatJPEG,
				Quality: quality,
			})
			if !errors.Is(err, toolkiterrors.ErrInvalidValue) {
				t.Fatalf("error = %v, want ErrInvalidValue", err)
			}
		})
	}
}

// --- sinks ----------------------------------------------------------------

func TestExportChartToSink_File(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))
	outPath := t.TempDir() + "/chart_export.png"

	err := converter.ExportChartToSink(src, datasource.FilePathSink(outPath), 0, 0,
		&converter.ChartExportOptions{Format: converter.ChartExportFormatPNG})
	if err != nil {
		t.Fatalf("ExportChartToSink: %v", err)
	}
	written, err := os.ReadFile(outPath)
	if err != nil {
		t.Fatalf("read exported file: %v", err)
	}
	if !isPNG(written) {
		t.Fatal("written file is not a PNG")
	}
}

func TestExportChartToFile(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))
	outPath := t.TempDir() + "/chart_file.pdf"

	err := converter.ExportChartToFile(src, outPath, 0, 0,
		&converter.ChartExportOptions{Format: converter.ChartExportFormatPDF})
	if err != nil {
		t.Fatalf("ExportChartToFile: %v", err)
	}
	written, err := os.ReadFile(outPath)
	if err != nil {
		t.Fatalf("read exported file: %v", err)
	}
	if !isPDF(written) {
		t.Fatal("file written with a .pdf name does not contain a PDF")
	}
}

// --- errors ---------------------------------------------------------------

func TestExportChartToBytes_InvalidChartIndex(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	_, err := converter.ExportChartToBytes(src, 0, 99, nil)
	if !errors.Is(err, toolkiterrors.ErrChartNotFound) {
		t.Fatalf("error = %v, want ErrChartNotFound", err)
	}
}

func TestExportChartToBytes_NegativeChartIndex(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	_, err := converter.ExportChartToBytes(src, 0, -1, nil)
	if !errors.Is(err, toolkiterrors.ErrChartNotFound) {
		t.Fatalf("error = %v, want ErrChartNotFound", err)
	}
}

// TestExportChartToBytes_InvalidSheetIndex checks the sheet lookup is guarded
// by a sentinel rather than surfacing a raw engine error.
func TestExportChartToBytes_InvalidSheetIndex(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	for _, sheetIndex := range []int{99, -1} {
		t.Run(fmt.Sprintf("sheet_%d", sheetIndex), func(t *testing.T) {
			_, err := converter.ExportChartToBytes(src, sheetIndex, 0, nil)
			if !errors.Is(err, toolkiterrors.ErrInvalidSheetID) {
				t.Fatalf("error = %v, want ErrInvalidSheetID", err)
			}
		})
	}
}

func TestExportChartToBytes_NilDataSource(t *testing.T) {
	_, err := converter.ExportChartToBytes(nil, 0, 0, nil)
	if !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Fatalf("error = %v, want ErrDataSourceNil", err)
	}
}

func TestExportChartToSink_NilDataSink(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	err := converter.ExportChartToSink(src, nil, 0, 0, nil)
	if !errors.Is(err, toolkiterrors.ErrDataSinkNil) {
		t.Fatalf("error = %v, want ErrDataSinkNil", err)
	}
}

func TestExportChartToBytes_UnsupportedFormat(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	_, err := converter.ExportChartToBytes(src, 0, 0, &converter.ChartExportOptions{
		Format: "unsupported",
	})
	if !errors.Is(err, toolkiterrors.ErrUnsupportedFormat) {
		t.Fatalf("error = %v, want ErrUnsupportedFormat", err)
	}
}

// --- selection and composition --------------------------------------------

func TestExportChartToBytes_MultipleCharts(t *testing.T) {
	seed, err := datasource.NewEmptyWorkbook()
	if err != nil {
		t.Fatalf("NewEmptyWorkbook: %v", err)
	}
	workbook, err := editor.EditSpreadsheet(
		seed,
		editor.InWorksheet(0,
			editor.SetCellValue(0, 0, "Data"),
			editor.SetCellValue(1, 0, 10),
			editor.SetCellValue(2, 0, 20),
			editor.AddChart(editor.ChartTypeColumn, "A1:A3", true, 0, 2, 12, 8,
				editor.WithChartTitle("Chart 1"),
			),
			editor.AddChart(editor.ChartTypePie, "A1:A3", true, 0, 10, 12, 16,
				editor.WithChartTitle("Chart 2"),
			),
		),
	)
	if err != nil {
		t.Fatalf("create multi-chart workbook: %v", err)
	}
	src := datasource.BytesSource(workbook)

	first, err := converter.ExportChartToBytes(src, 0, 0, nil)
	if err != nil {
		t.Fatalf("export chart 0: %v", err)
	}
	second, err := converter.ExportChartToBytes(src, 0, 1, nil)
	if err != nil {
		t.Fatalf("export chart 1: %v", err)
	}
	if !isPNG(first) || !isPNG(second) {
		t.Fatal("both exports should be PNG")
	}
	if bytes.Equal(first, second) {
		t.Error("chart 0 and chart 1 exported identical bytes; index is not selecting the chart")
	}
}

func TestExportChartToBytes_AfterModification(t *testing.T) {
	modified, err := editor.EditSpreadsheet(
		datasource.BytesSource(getChartExportFixture(t)),
		editor.InWorksheet(0,
			editor.InChart(0, editor.WithChartTitle("Modified Chart Title")),
		),
	)
	if err != nil {
		t.Fatalf("modify chart: %v", err)
	}

	data, err := converter.ExportChartToBytes(datasource.BytesSource(modified), 0, 0, nil)
	if err != nil {
		t.Fatalf("export modified chart: %v", err)
	}
	if !isPNG(data) {
		t.Fatal("exported modified chart is not a PNG")
	}
}

func TestExportChartIntegration(t *testing.T) {
	src := datasource.BytesSource(getChartExportFixture(t))

	count, err := query.ChartCount(src)
	if err != nil {
		t.Fatalf("query chart count: %v", err)
	}
	if count != 1 {
		t.Fatalf("expected 1 chart, got %d", count)
	}

	info, err := query.ChartInfo(src, 0)
	if err != nil {
		t.Fatalf("query chart info: %v", err)
	}
	if info.Title != "Test Chart" {
		t.Errorf("title = %q, want %q", info.Title, "Test Chart")
	}

	data, err := converter.ExportChartToBytes(src, 0, 0, nil)
	if err != nil {
		t.Fatalf("export chart: %v", err)
	}
	if !isPNG(data) {
		t.Fatal("exported chart is not a PNG")
	}
}

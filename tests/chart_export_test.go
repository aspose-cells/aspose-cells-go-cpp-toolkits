package tests

import (
	"bytes"
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

// getChartExportFixture returns a workbook with a chart for testing.
func getChartExportFixture(t *testing.T) []byte {
	t.Helper()
	chartExportFixtureOnce.Do(func() {
		// Create a workbook with sample data and a chart
		src, err := datasource.NewEmptyWorkbook()
		if err != nil {
			chartExportFixtureErr = err
			return
		}
		out, err := editor.EditSpreadsheet(
			src,
			editor.InWorksheet(0,
				// Add sample data
				editor.SetCellValue(0, 0, "Category"),
				editor.SetCellValue(0, 1, "Value"),
				editor.SetCellValue(1, 0, "A"),
				editor.SetCellValue(1, 1, 10),
				editor.SetCellValue(2, 0, "B"),
				editor.SetCellValue(2, 1, 20),
				editor.SetCellValue(3, 0, "C"),
				editor.SetCellValue(3, 1, 30),
				// Add a chart
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

func TestExportChartToBytes_PNG(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	opts := &converter.ChartExportOptions{
		Format: converter.ChartExportFormatPNG,
	}

	data, err := converter.ExportChartToBytes(dataSource, 0, 0, opts)
	if err != nil {
		t.Fatalf("ExportChartToBytes PNG: %v", err)
	}

	if len(data) == 0 {
		t.Fatal("Exported PNG data is empty")
	}

	// PNG files start with specific magic bytes
	if len(data) < 8 || data[0] != 0x89 || data[1] != 0x50 || data[2] != 0x4E || data[3] != 0x47 {
		t.Error("Exported data does not appear to be PNG format")
	}
}

func TestExportChartToBytes_JPEG(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	opts := &converter.ChartExportOptions{
		Format:  converter.ChartExportFormatJPEG,
		Quality: 85,
	}

	data, err := converter.ExportChartToBytes(dataSource, 0, 0, opts)
	if err != nil {
		t.Fatalf("ExportChartToBytes JPEG: %v", err)
	}

	if len(data) == 0 {
		t.Fatal("Exported JPEG data is empty")
	}

	// JPEG files start with specific magic bytes
	if len(data) < 2 || data[0] != 0xFF || data[1] != 0xD8 {
		t.Error("Exported data does not appear to be JPEG format")
	}
}

func TestExportChartToBytes_SVG(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	opts := &converter.ChartExportOptions{
		Format: converter.ChartExportFormatSVG,
	}

	data, err := converter.ExportChartToBytes(dataSource, 0, 0, opts)
	if err != nil {
		t.Fatalf("ExportChartToBytes SVG: %v", err)
	}

	if len(data) == 0 {
		t.Fatal("Exported SVG data is empty")
	}

	// SVG files are XML and should contain "<svg"
	svgContent := string(data)
	if !bytes.Contains(data, []byte("<svg")) && !bytes.Contains(data, []byte("<?xml")) {
		end := 100
		if len(svgContent) < 100 {
			end = len(svgContent)
		}
		t.Errorf("Exported data does not appear to be SVG format, got: %s", svgContent[:end])
	}
}

func TestExportChartToBytes_PDF(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	opts := &converter.ChartExportOptions{
		Format: converter.ChartExportFormatPDF,
	}

	data, err := converter.ExportChartToBytes(dataSource, 0, 0, opts)
	if err != nil {
		t.Fatalf("ExportChartToBytes PDF (EMF): %v", err)
	}

	if len(data) == 0 {
		t.Fatal("Exported PDF (EMF) data is empty")
	}
}

func TestExportChartToBytes_WithDimensions(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	opts := &converter.ChartExportOptions{
		Format: converter.ChartExportFormatPNG,
		Width:  800,
		Height: 600,
	}

	data, err := converter.ExportChartToBytes(dataSource, 0, 0, opts)
	if err != nil {
		t.Fatalf("ExportChartToBytes with dimensions: %v", err)
	}

	if len(data) == 0 {
		t.Fatal("Exported PNG data is empty")
	}
}

func TestExportChartToBytes_WidthOnly(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	// Test with Width > 0, Height = 0 (should use default height)
	opts := &converter.ChartExportOptions{
		Format: converter.ChartExportFormatPNG,
		Width:  1000,
		Height: 0,
	}

	data, err := converter.ExportChartToBytes(dataSource, 0, 0, opts)
	if err != nil {
		t.Fatalf("ExportChartToBytes with width only: %v", err)
	}

	if len(data) == 0 {
		t.Fatal("Exported PNG data is empty")
	}
}

func TestExportChartToBytes_HeightOnly(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	// Test with Width = 0, Height > 0 (should use default width)
	opts := &converter.ChartExportOptions{
		Format: converter.ChartExportFormatPNG,
		Width:  0,
		Height: 800,
	}

	data, err := converter.ExportChartToBytes(dataSource, 0, 0, opts)
	if err != nil {
		t.Fatalf("ExportChartToBytes with height only: %v", err)
	}

	if len(data) == 0 {
		t.Fatal("Exported PNG data is empty")
	}
}

func TestExportChartToBytes_JPEGQualityEdgeCases(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	// Test various JPEG quality values
	qualityLevels := []int{1, 50, 95, 100}

	for _, quality := range qualityLevels {
		t.Run(fmt.Sprintf("Quality_%d", quality), func(t *testing.T) {
			opts := &converter.ChartExportOptions{
				Format:  converter.ChartExportFormatJPEG,
				Quality: quality,
			}

			data, err := converter.ExportChartToBytes(dataSource, 0, 0, opts)
			if err != nil {
				t.Fatalf("ExportChartToBytes JPEG with quality %d: %v", quality, err)
			}

			if len(data) == 0 {
				t.Fatalf("Exported JPEG data is empty for quality %d", quality)
			}

			// Verify it's still JPEG format
			if len(data) < 2 || data[0] != 0xFF || data[1] != 0xD8 {
				t.Errorf("Quality %d: exported data is not JPEG format", quality)
			}
		})
	}
}

func TestExportChartToSink_File(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	tmpFile := t.TempDir() + "/chart_export.png"
	sink := datasource.FilePathSink(tmpFile)

	opts := &converter.ChartExportOptions{
		Format: converter.ChartExportFormatPNG,
	}

	err := converter.ExportChartToSink(dataSource, sink, 0, 0, opts)
	if err != nil {
		t.Fatalf("ExportChartToSink: %v", err)
	}

	// Verify file was created and has content
	fileData, err := os.ReadFile(tmpFile)
	if err != nil {
		t.Fatalf("Read exported file: %v", err)
	}

	if len(fileData) == 0 {
		t.Fatal("Exported file is empty")
	}
}

func TestExportChartToFile(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	tmpFile := t.TempDir() + "/chart_export_file.png"

	opts := &converter.ChartExportOptions{
		Format: converter.ChartExportFormatPNG,
	}

	err := converter.ExportChartToFile(dataSource, tmpFile, 0, 0, opts)
	if err != nil {
		t.Fatalf("ExportChartToFile: %v", err)
	}

	// Verify file was created
	fileData, err := os.ReadFile(tmpFile)
	if err != nil {
		t.Fatalf("Read exported file: %v", err)
	}

	if len(fileData) == 0 {
		t.Fatal("Exported file is empty")
	}
}

func TestExportChartToBytes_InvalidChartIndex(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	opts := &converter.ChartExportOptions{
		Format: converter.ChartExportFormatPNG,
	}

	_, err := converter.ExportChartToBytes(dataSource, 0, 99, opts)
	if err == nil {
		t.Fatal("Expected error for invalid chart index")
	}

	if !errors.Is(err, toolkiterrors.ErrChartNotFound) {
		t.Errorf("Expected ErrChartNotFound, got: %v", err)
	}
}

func TestExportChartToBytes_InvalidSheetIndex(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	opts := &converter.ChartExportOptions{
		Format: converter.ChartExportFormatPNG,
	}

	_, err := converter.ExportChartToBytes(dataSource, 99, 0, opts)
	if err == nil {
		t.Fatal("Expected error for invalid sheet index")
	}
}

func TestExportChartToBytes_NilDataSource(t *testing.T) {
	opts := &converter.ChartExportOptions{
		Format: converter.ChartExportFormatPNG,
	}

	_, err := converter.ExportChartToBytes(nil, 0, 0, opts)
	if err == nil {
		t.Fatal("Expected error for nil data source")
	}

	if !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Errorf("Expected ErrDataSourceNil, got: %v", err)
	}
}

func TestExportChartToSink_NilDataSink(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	opts := &converter.ChartExportOptions{
		Format: converter.ChartExportFormatPNG,
	}

	err := converter.ExportChartToSink(dataSource, nil, 0, 0, opts)
	if err == nil {
		t.Fatal("Expected error for nil data sink")
	}

	if !errors.Is(err, toolkiterrors.ErrDataSinkNil) {
		t.Errorf("Expected ErrDataSinkNil, got: %v", err)
	}
}

func TestExportChartToBytes_UnsupportedFormat(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	opts := &converter.ChartExportOptions{
		Format: "unsupported",
	}

	_, err := converter.ExportChartToBytes(dataSource, 0, 0, opts)
	if err == nil {
		t.Fatal("Expected error for unsupported format")
	}

	if !errors.Is(err, toolkiterrors.ErrUnsupportedFormat) {
		t.Errorf("Expected ErrUnsupportedFormat, got: %v", err)
	}
}

func TestExportChartToBytes_DefaultOptions(t *testing.T) {
	src := getChartExportFixture(t)
	dataSource := datasource.BytesSource(src)

	// Pass nil options - should default to PNG
	data, err := converter.ExportChartToBytes(dataSource, 0, 0, nil)
	if err != nil {
		t.Fatalf("ExportChartToBytes with nil options: %v", err)
	}

	if len(data) == 0 {
		t.Fatal("Exported data is empty")
	}

	// Should be PNG format
	if len(data) < 8 || data[0] != 0x89 || data[1] != 0x50 {
		t.Error("Default format should be PNG")
	}
}

func TestExportChartToBytes_MultipleCharts(t *testing.T) {
	// Create a workbook with multiple charts
	src, err := datasource.NewEmptyWorkbook()
	if err != nil {
		t.Fatalf("NewEmptyWorkbook: %v", err)
	}
	out, err := editor.EditSpreadsheet(
		src,
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
		t.Fatalf("Create multi-chart fixture: %v", err)
	}

	dataSource := datasource.BytesSource(out)

	// Export first chart
	data1, err := converter.ExportChartToBytes(dataSource, 0, 0, nil)
	if err != nil {
		t.Fatalf("Export first chart: %v", err)
	}
	if len(data1) == 0 {
		t.Fatal("First chart export is empty")
	}

	// Export second chart
	data2, err := converter.ExportChartToBytes(dataSource, 0, 1, nil)
	if err != nil {
		t.Fatalf("Export second chart: %v", err)
	}
	if len(data2) == 0 {
		t.Fatal("Second chart export is empty")
	}
}

func TestExportChartToBytes_AfterModification(t *testing.T) {
	// Create initial workbook with chart
	src := getChartExportFixture(t)

	// Modify the chart title
	modified, err := editor.EditSpreadsheet(
		datasource.BytesSource(src),
		editor.InWorksheet(0,
			editor.InChart(0,
				editor.WithChartTitle("Modified Chart Title"),
			),
		),
	)
	if err != nil {
		t.Fatalf("Modify chart: %v", err)
	}

	// Export the modified chart
	data, err := converter.ExportChartToBytes(datasource.BytesSource(modified), 0, 0, nil)
	if err != nil {
		t.Fatalf("Export modified chart: %v", err)
	}

	if len(data) == 0 {
		t.Fatal("Exported modified chart is empty")
	}
}

func TestExportChartIntegration(t *testing.T) {
	// Integration test: create workbook, export chart, verify chart info matches
	src := getChartExportFixture(t)

	// Verify chart exists
	count, err := query.ChartCount(datasource.BytesSource(src))
	if err != nil {
		t.Fatalf("Query chart count: %v", err)
	}
	if count != 1 {
		t.Fatalf("Expected 1 chart, got %d", count)
	}

	// Export the chart
	data, err := converter.ExportChartToBytes(datasource.BytesSource(src), 0, 0, nil)
	if err != nil {
		t.Fatalf("Export chart: %v", err)
	}

	if len(data) == 0 {
		t.Fatal("Exported chart is empty")
	}

	// Get chart info
	info, err := query.ChartInfo(datasource.BytesSource(src), 0)
	if err != nil {
		t.Fatalf("Query chart info: %v", err)
	}

	if info.Title != "Test Chart" {
		t.Errorf("Expected title 'Test Chart', got '%s'", info.Title)
	}
}

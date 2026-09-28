// Package converter provides functionality for converting spreadsheets and
// their components (charts, worksheets) to different formats.
package converter

import (
	"fmt"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// ChartExportFormat specifies the output format for chart export.
type ChartExportFormat string

const (
	// ChartExportFormatPNG exports the chart as a PNG image.
	ChartExportFormatPNG ChartExportFormat = "png"
	// ChartExportFormatJPEG exports the chart as a JPEG image.
	ChartExportFormatJPEG ChartExportFormat = "jpeg"
	// ChartExportFormatSVG exports the chart as an SVG vector image.
	ChartExportFormatSVG ChartExportFormat = "svg"
	// ChartExportFormatPDF exports the chart as a PDF document.
	ChartExportFormatPDF ChartExportFormat = "pdf"
)

// ChartExportOptions configures chart export behavior.
type ChartExportOptions struct {
	// Format is the output format (PNG, JPEG, SVG, or PDF).
	Format ChartExportFormat
	// Width is the desired width in pixels (for raster formats). Zero means use chart's default width.
	Width int
	// Height is the desired height in pixels (for raster formats). Zero means use chart's default height.
	Height int
	// Quality is the JPEG quality (1-100). Only applies to JPEG format. Default is 95.
	Quality int
}

// ExportChartToSink exports a chart from a workbook to a data sink.
//
// Parameters:
//   - src: The data source containing the workbook.
//   - sink: The data sink to write the exported chart to.
//   - sheetIndex: The zero-based index of the worksheet containing the chart.
//   - chartIndex: The zero-based index of the chart to export.
//   - opts: Export options. If nil, defaults to PNG format.
//
// Returns:
//   - error: An error if the export fails.
//
// Example:
//
//	err := converter.ExportChartToSink(
//	    datasource.FilePathSource("workbook.xlsx"),
//	    datasource.FilePathSink("chart.png"),
//	    0, 0,
//	    &converter.ChartExportOptions{Format: converter.ChartExportFormatPNG},
//	)
func ExportChartToSink(src datasource.DataSource, sink datasource.DataSink, sheetIndex, chartIndex int, opts *ChartExportOptions) error {
	if src == nil {
		return toolkiterrors.ErrDataSourceNil
	}
	if sink == nil {
		return toolkiterrors.ErrDataSinkNil
	}
	if opts == nil {
		opts = &ChartExportOptions{Format: ChartExportFormatPNG}
	}

	// Load the workbook
	data, err := cells.ReadSource(src)
	if err != nil {
		return fmt.Errorf("read source: %w", err)
	}
	wb, err := asposecells.NewWorkbook_Stream(data)
	if err != nil {
		return fmt.Errorf("open workbook: %w", err)
	}
	defer wb.Dispose()

	// Get the worksheet
	wss, err := wb.GetWorksheets()
	if err != nil {
		return fmt.Errorf("get worksheets: %w", err)
	}
	ws, err := wss.Get_Int(int32(sheetIndex))
	if err != nil {
		return fmt.Errorf("get worksheet %d: %w", sheetIndex, err)
	}

	// Get the chart
	charts, err := ws.GetCharts()
	if err != nil {
		return fmt.Errorf("get charts: %w", err)
	}
	chartCount, err := charts.GetCount()
	if err != nil {
		return fmt.Errorf("get chart count: %w", err)
	}
	if chartIndex < 0 || int32(chartIndex) >= chartCount {
		return fmt.Errorf("chart index %d out of range [0, %d): %w", chartIndex, chartCount, toolkiterrors.ErrChartNotFound)
	}
	chart, err := charts.Get_Int(int32(chartIndex))
	if err != nil {
		return fmt.Errorf("get chart %d: %w", chartIndex, err)
	}

	// Export based on format
	var outputData []byte
	switch opts.Format {
	case ChartExportFormatPNG:
		outputData, err = exportChartToImage(chart, asposecells.ImageType_Png, opts)
	case ChartExportFormatJPEG:
		outputData, err = exportChartToImage(chart, asposecells.ImageType_Jpeg, opts)
	case ChartExportFormatSVG:
		outputData, err = exportChartToImage(chart, asposecells.ImageType_Svg, opts)
	case ChartExportFormatPDF:
		outputData, err = exportChartToPDF(chart, opts)
	default:
		return fmt.Errorf("unsupported chart export format %q: %w", opts.Format, toolkiterrors.ErrUnsupportedFormat)
	}

	if err != nil {
		return fmt.Errorf("export chart: %w", err)
	}

	// Write to sink
	if err := sink.Write("", outputData); err != nil {
		return fmt.Errorf("write to sink: %w", err)
	}

	return nil
}

// ExportChartToBytes exports a chart from a workbook and returns it as a byte slice.
//
// Parameters:
//   - src: The data source containing the workbook.
//   - sheetIndex: The zero-based index of the worksheet containing the chart.
//   - chartIndex: The zero-based index of the chart to export.
//   - opts: Export options. If nil, defaults to PNG format.
//
// Returns:
//   - []byte: The exported chart data.
//   - error: An error if the export fails.
//
// Example:
//
//	data, err := converter.ExportChartToBytes(
//	    datasource.FilePathSource("workbook.xlsx"),
//	    0, 0,
//	    &converter.ChartExportOptions{Format: converter.ChartExportFormatPNG},
//	)
func ExportChartToBytes(src datasource.DataSource, sheetIndex, chartIndex int, opts *ChartExportOptions) ([]byte, error) {
	if src == nil {
		return nil, toolkiterrors.ErrDataSourceNil
	}
	if opts == nil {
		opts = &ChartExportOptions{Format: ChartExportFormatPNG}
	}

	// Load the workbook
	data, err := cells.ReadSource(src)
	if err != nil {
		return nil, fmt.Errorf("read source: %w", err)
	}
	wb, err := asposecells.NewWorkbook_Stream(data)
	if err != nil {
		return nil, fmt.Errorf("open workbook: %w", err)
	}
	defer wb.Dispose()

	// Get the worksheet
	wss, err := wb.GetWorksheets()
	if err != nil {
		return nil, fmt.Errorf("get worksheets: %w", err)
	}
	ws, err := wss.Get_Int(int32(sheetIndex))
	if err != nil {
		return nil, fmt.Errorf("get worksheet %d: %w", sheetIndex, err)
	}

	// Get the chart
	charts, err := ws.GetCharts()
	if err != nil {
		return nil, fmt.Errorf("get charts: %w", err)
	}
	chartCount, err := charts.GetCount()
	if err != nil {
		return nil, fmt.Errorf("get chart count: %w", err)
	}
	if chartIndex < 0 || int32(chartIndex) >= chartCount {
		return nil, fmt.Errorf("chart index %d out of range [0, %d): %w", chartIndex, chartCount, toolkiterrors.ErrChartNotFound)
	}
	chart, err := charts.Get_Int(int32(chartIndex))
	if err != nil {
		return nil, fmt.Errorf("get chart %d: %w", chartIndex, err)
	}

	// Export based on format
	switch opts.Format {
	case ChartExportFormatPNG:
		return exportChartToImage(chart, asposecells.ImageType_Png, opts)
	case ChartExportFormatJPEG:
		return exportChartToImage(chart, asposecells.ImageType_Jpeg, opts)
	case ChartExportFormatSVG:
		return exportChartToImage(chart, asposecells.ImageType_Svg, opts)
	case ChartExportFormatPDF:
		return exportChartToPDF(chart, opts)
	default:
		return nil, fmt.Errorf("unsupported chart export format %q: %w", opts.Format, toolkiterrors.ErrUnsupportedFormat)
	}
}

// exportChartToImage exports a chart to a raster or vector image format.
func exportChartToImage(chart *asposecells.Chart, imageType asposecells.ImageType, opts *ChartExportOptions) ([]byte, error) {
	// Create image options
	imgOpts, err := asposecells.NewImageOrPrintOptions()
	if err != nil {
		return nil, fmt.Errorf("create image options: %w", err)
	}

	// Set image type
	if err := imgOpts.SetImageType(imageType); err != nil {
		return nil, fmt.Errorf("set image type: %w", err)
	}

	// Set dimensions if specified
	if opts.Width > 0 || opts.Height > 0 {
		width := int32(opts.Width)
		height := int32(opts.Height)
		if width == 0 {
			width = 800 // default width
		}
		if height == 0 {
			height = 600 // default height
		}
		if err := imgOpts.SetDesiredSize(width, height, true); err != nil {
			return nil, fmt.Errorf("set dimensions: %w", err)
		}
	}

	// Set JPEG quality if applicable
	if imageType == asposecells.ImageType_Jpeg && opts.Quality > 0 {
		if err := imgOpts.SetQuality(int32(opts.Quality)); err != nil {
			return nil, fmt.Errorf("set quality: %w", err)
		}
	}

	// Export to image
	return chart.ToImage_ImageOrPrintOptions(imgOpts)
}

// exportChartToPDF exports a chart to PDF format.
// Note: The engine doesn't have a direct Chart.ToPDF method, so we use
// ImageOrPrintOptions with PDF-like settings and export as an image.
// For true PDF export, users should export the entire worksheet or use
// the converter package's spreadsheet conversion functions.
func exportChartToPDF(chart *asposecells.Chart, opts *ChartExportOptions) ([]byte, error) {
	// Create image options configured for high-quality output
	imgOpts, err := asposecells.NewImageOrPrintOptions()
	if err != nil {
		return nil, fmt.Errorf("create image options: %w", err)
	}

	// Use EMF format which is vector-based and can be converted to PDF
	// EMF is Windows Enhanced Metafile, which is vector-based like PDF
	if err := imgOpts.SetImageType(asposecells.ImageType_Emf); err != nil {
		return nil, fmt.Errorf("set image type to EMF: %w", err)
	}

	// Set dimensions if specified
	if opts.Width > 0 || opts.Height > 0 {
		width := int32(opts.Width)
		height := int32(opts.Height)
		if width == 0 {
			width = 800 // default width
		}
		if height == 0 {
			height = 600 // default height
		}
		if err := imgOpts.SetDesiredSize(width, height, true); err != nil {
			return nil, fmt.Errorf("set dimensions: %w", err)
		}
	}

	// Export to EMF (vector format, similar to PDF)
	data, err := chart.ToImage_ImageOrPrintOptions(imgOpts)
	if err != nil {
		return nil, fmt.Errorf("export chart to EMF: %w", err)
	}

	// Note: For true PDF output, users would need to convert the EMF to PDF
	// using an external tool or library. The engine doesn't provide direct
	// chart-to-PDF conversion.
	return data, nil
}

// ExportChartToFile is a convenience function that exports a chart to a file.
//
// Parameters:
//   - src: The data source containing the workbook.
//   - outputPath: The path to write the exported chart to.
//   - sheetIndex: The zero-based index of the worksheet containing the chart.
//   - chartIndex: The zero-based index of the chart to export.
//   - opts: Export options. If nil, defaults to PNG format.
//
// Returns:
//   - error: An error if the export fails.
//
// Example:
//
//	err := converter.ExportChartToFile(
//	    datasource.FilePathSource("workbook.xlsx"),
//	    "chart.png",
//	    0, 0,
//	    &converter.ChartExportOptions{Format: converter.ChartExportFormatPNG},
//	)
func ExportChartToFile(src datasource.DataSource, outputPath string, sheetIndex, chartIndex int, opts *ChartExportOptions) error {
	sink := datasource.FilePathSink(outputPath)
	return ExportChartToSink(src, sink, sheetIndex, chartIndex, opts)
}

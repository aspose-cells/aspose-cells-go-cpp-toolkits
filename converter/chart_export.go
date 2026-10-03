package converter

import (
	"encoding/binary"
	"fmt"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pdf"
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
	// ChartExportFormatPDF exports the chart as a real PDF document: the chart
	// is rendered to an image, embedded in a single-sheet workbook, and that
	// workbook is saved as PDF. The PDF is therefore a raster image of the
	// chart on a page sized to it, not vector artwork.
	ChartExportFormatPDF ChartExportFormat = "pdf"
)

// Default row height (points-to-pixels at 96 DPI for Excel's 15pt default row)
// and column width (Excel's 8.43-character default) used to translate a pixel
// size into the cell rectangle a PDF page is sized to.
const (
	defaultRowHeightPx   = 20
	defaultColumnWidthPx = 64
)

// ChartExportOptions configures chart export behavior.
type ChartExportOptions struct {
	// Format is the output format (PNG, JPEG, SVG, or PDF).
	Format ChartExportFormat
	// Width and Height are the exact output size in pixels. Either both are
	// set to a positive number, or both are left at zero to use the chart's own
	// size. Setting exactly one is an error (ErrInvalidChartSize), because the
	// engine cannot size an image from one dimension alone.
	Width  int
	Height int
	// Quality is the JPEG quality, 1-100. Only applies to the JPEG format; zero
	// leaves the engine's default. A value outside 1-100 is ErrInvalidValue.
	Quality int
}

// ExportChartToSink exports a chart from a workbook to a data sink.
//
// Parameters:
//   - src: The data source containing the workbook.
//   - sink: The data sink to write the exported chart to.
//   - sheetIndex: The zero-based index of the worksheet containing the chart.
//   - chartIndex: The zero-based index of the chart to export.
//   - opts: Export options. If nil, defaults to PNG format at the chart's own size.
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
	data, err := exportChart(src, sheetIndex, chartIndex, opts)
	if err != nil {
		return err
	}
	if err := sink.Write("", data); err != nil {
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
//   - opts: Export options. If nil, defaults to PNG format at the chart's own size.
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
	return exportChart(src, sheetIndex, chartIndex, opts)
}

// ExportChartToFile is a convenience function that exports a chart to a file.
//
// Parameters:
//   - src: The data source containing the workbook.
//   - outputPath: The path to write the exported chart to.
//   - sheetIndex: The zero-based index of the worksheet containing the chart.
//   - chartIndex: The zero-based index of the chart to export.
//   - opts: Export options. If nil, defaults to PNG format at the chart's own size.
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
	return ExportChartToSink(src, datasource.FilePathSink(outputPath), sheetIndex, chartIndex, opts)
}

// exportChart is the single export path. Both public byte-returning entry
// points route through it, so the workbook load, the sheet lookup, the chart
// lookup, and the format dispatch exist exactly once.
func exportChart(src datasource.DataSource, sheetIndex, chartIndex int, opts *ChartExportOptions) ([]byte, error) {
	engine.LockEngine()
	defer engine.UnlockEngine()
	if opts == nil {
		opts = &ChartExportOptions{Format: ChartExportFormatPNG}
	}
	width, height, err := resolveSize(opts)
	if err != nil {
		return nil, err
	}
	if err := validateQuality(opts); err != nil {
		return nil, err
	}

	chart, workbook, err := resolveChart(src, sheetIndex, chartIndex)
	if err != nil {
		return nil, err
	}
	defer engine.CloseWorkbook(workbook)

	switch opts.Format {
	case ChartExportFormatPNG:
		return renderChart(chart, asposecells.ImageType_Png, opts, width, height)
	case ChartExportFormatJPEG:
		return renderChart(chart, asposecells.ImageType_Jpeg, opts, width, height)
	case ChartExportFormatSVG:
		return renderChart(chart, asposecells.ImageType_Svg, opts, width, height)
	case ChartExportFormatPDF:
		return renderChartToPDF(chart, opts, width, height)
	default:
		return nil, fmt.Errorf("unsupported chart export format %q: %w", opts.Format, toolkiterrors.ErrUnsupportedFormat)
	}
}

// resolveChart loads the workbook from src and returns the chart at
// (sheetIndex, chartIndex) together with the workbook that owns it. The caller
// owns the workbook and must Dispose it.
//
// Both lookups run through the guarded internal helpers, which range-check the
// index against the collection before touching the engine: the binding returns
// a dangling handle for an out-of-range index, and using one crashes the
// process. This is the only place either lookup happens.
func resolveChart(src datasource.DataSource, sheetIndex, chartIndex int) (*asposecells.Chart, *asposecells.Workbook, error) {
	workbook, err := cells.GetWorkbookWithDataSource(src)
	if err != nil {
		return nil, nil, fmt.Errorf("open workbook: %w", err)
	}
	worksheet, err := cells.SheetByIndex(workbook, sheetIndex)
	if err != nil {
		engine.CloseWorkbook(workbook)
		return nil, nil, err
	}
	chart, err := cells.Chart(worksheet, chartIndex)
	if err != nil {
		engine.CloseWorkbook(workbook)
		return nil, nil, err
	}
	return chart, workbook, nil
}

// resolveSize validates the requested output size and reports whether it was
// set. Following the engine's SetDesiredSize contract, a size is either fully
// specified (both dimensions positive) or not specified at all; a half-specified
// size cannot be honored and is an error rather than a guess.
func resolveSize(opts *ChartExportOptions) (width, height int, err error) {
	if opts.Width < 0 || opts.Height < 0 {
		return 0, 0, fmt.Errorf("chart export size %dx%d is negative: %w",
			opts.Width, opts.Height, toolkiterrors.ErrInvalidChartSize)
	}
	if (opts.Width == 0) != (opts.Height == 0) {
		return 0, 0, fmt.Errorf("chart export size %dx%d sets only one dimension; set both or neither: %w",
			opts.Width, opts.Height, toolkiterrors.ErrInvalidChartSize)
	}
	return opts.Width, opts.Height, nil
}

// validateQuality rejects a JPEG quality outside the engine's 1-100 range
// instead of letting the engine silently clamp it. Zero means "unset".
func validateQuality(opts *ChartExportOptions) error {
	if opts.Quality != 0 && (opts.Quality < 1 || opts.Quality > 100) {
		return fmt.Errorf("JPEG quality %d is outside 1-100: %w", opts.Quality, toolkiterrors.ErrInvalidValue)
	}
	return nil
}

// renderChart renders the chart to a raster or vector image. When width and
// height are positive the output is exactly that size; when both are zero the
// chart's own size is used.
func renderChart(chart *asposecells.Chart, imageType asposecells.ImageType, opts *ChartExportOptions, width, height int) ([]byte, error) {
	imgOpts, err := asposecells.NewImageOrPrintOptions()
	if err != nil {
		return nil, fmt.Errorf("create image options: %w", err)
	}
	if err := imgOpts.SetImageType(imageType); err != nil {
		return nil, fmt.Errorf("set image type: %w", err)
	}
	if width > 0 && height > 0 {
		// keepAspectRatio=false: the caller asked for exact pixels, so the
		// output must be exactly width x height rather than a letterboxed fit.
		if err := imgOpts.SetDesiredSize(int32(width), int32(height), false); err != nil {
			return nil, fmt.Errorf("set dimensions: %w", err)
		}
	}
	if imageType == asposecells.ImageType_Jpeg && opts.Quality > 0 {
		if err := imgOpts.SetQuality(int32(opts.Quality)); err != nil {
			return nil, fmt.Errorf("set quality: %w", err)
		}
	}
	data, err := chart.ToImage_ImageOrPrintOptions(imgOpts)
	if err != nil {
		return nil, fmt.Errorf("render chart: %w", err)
	}
	return data, nil
}

// renderChartToPDF produces a real PDF containing the chart: the chart is
// rendered to a PNG, embedded into a fresh one-sheet workbook on a sheet sized
// to the image, and that workbook is converted with the PDF save option. The
// result opens in any PDF reader.
//
// The chart appears as a raster image on the page; this is the toolkit's only
// chart-to-PDF route, because the engine exposes no vector chart-to-PDF call.
func renderChartToPDF(chart *asposecells.Chart, opts *ChartExportOptions, width, height int) ([]byte, error) {
	image, err := renderChart(chart, asposecells.ImageType_Png, opts, width, height)
	if err != nil {
		return nil, err
	}
	pxWidth, pxHeight, err := pngSize(image)
	if err != nil {
		return nil, err
	}
	workbook, err := chartImageWorkbook(image, pxWidth, pxHeight)
	if err != nil {
		return nil, err
	}
	defer engine.CloseWorkbook(workbook)

	sheetBytes, err := cells.WorkbookToByteData(workbook)
	if err != nil {
		return nil, fmt.Errorf("serialize chart page: %w", err)
	}
	out, err := pdf.New(pdf.WithOnePagePerSheet(true)).Apply(sheetBytes)
	if err != nil {
		return nil, fmt.Errorf("convert chart page to PDF: %w", err)
	}
	return out, nil
}

// chartImageWorkbook builds a single-sheet workbook holding image on a sheet
// whose print area is exactly the image's cell span, so the PDF page covers the
// chart and nothing else. The caller owns the returned workbook.
func chartImageWorkbook(image []byte, pxWidth, pxHeight int) (*asposecells.Workbook, error) {
	workbook, err := engine.NewWorkbook()
	if err != nil {
		return nil, fmt.Errorf("create chart page workbook: %w", err)
	}
	worksheet, err := cells.SheetByIndex(workbook, 0)
	if err != nil {
		engine.CloseWorkbook(workbook)
		return nil, err
	}
	rows := ceilDiv(pxHeight, defaultRowHeightPx)
	columns := ceilDiv(pxWidth, defaultColumnWidthPx)
	if err := cells.AddPictureAt(worksheet, 0, 0, rows-1, columns-1, image); err != nil {
		engine.CloseWorkbook(workbook)
		return nil, err
	}
	pageSetup, err := worksheet.GetPageSetup()
	if err != nil {
		engine.CloseWorkbook(workbook)
		return nil, fmt.Errorf("get page setup: %w", err)
	}
	area := cells.Area{
		Start: cells.CellRef{Row: 0, Col: 0},
		End:   cells.CellRef{Row: rows - 1, Col: columns - 1},
	}
	if err := pageSetup.SetPrintArea(area.String()); err != nil {
		engine.CloseWorkbook(workbook)
		return nil, fmt.Errorf("set print area: %w", err)
	}
	if err := pageSetup.SetFitToPages(1, 1); err != nil {
		engine.CloseWorkbook(workbook)
		return nil, fmt.Errorf("fit chart page: %w", err)
	}
	return workbook, nil
}

// pngSize reads the pixel dimensions from a PNG's IHDR chunk, which is always
// the first chunk: 8 signature bytes, 4 length bytes, 4 type bytes, then the
// 4-byte big-endian width and height.
func pngSize(data []byte) (width, height int, err error) {
	const headerLen = 24
	if len(data) < headerLen {
		return 0, 0, fmt.Errorf("rendered image is %d bytes, too short to be a PNG: %w",
			len(data), toolkiterrors.ErrPictureAddFailed)
	}
	if string(data[12:16]) != "IHDR" {
		return 0, 0, fmt.Errorf("rendered image has no PNG IHDR header: %w", toolkiterrors.ErrPictureAddFailed)
	}
	width = int(binary.BigEndian.Uint32(data[16:20]))
	height = int(binary.BigEndian.Uint32(data[20:24]))
	if width <= 0 || height <= 0 {
		return 0, 0, fmt.Errorf("rendered image reports size %dx%d: %w",
			width, height, toolkiterrors.ErrPictureAddFailed)
	}
	return width, height, nil
}

// ceilDiv divides a by b, rounding up, and never returns less than 1.
func ceilDiv(a, b int) int {
	if a <= b {
		return 1
	}
	return (a + b - 1) / b
}

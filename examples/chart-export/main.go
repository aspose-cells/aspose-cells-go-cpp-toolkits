// Example: chart-export
//
// Demonstrates exporting charts from a workbook to various image formats
// (PNG, JPEG, SVG) using the converter package.
package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
)

func main() {
	// Create output directory
	outDir := filepath.Join("examples", "chart-export", "out")
	if err := os.MkdirAll(outDir, 0755); err != nil {
		log.Fatalf("create output directory: %v", err)
	}

	// Create a workbook with sample data and a chart
	src, err := datasource.NewEmptyWorkbook()
	if err != nil {
		log.Fatalf("create empty workbook: %v", err)
	}

	workbookData, err := editor.EditSpreadsheet(
		src,
		editor.InWorksheet(0,
			// Add header
			editor.SetCellValue(0, 0, "Quarter"),
			editor.SetCellValue(0, 1, "Sales"),
			editor.SetCellValue(0, 2, "Profit"),
			// Add data
			editor.SetCellValue(1, 0, "Q1"),
			editor.SetCellValue(1, 1, 100),
			editor.SetCellValue(1, 2, 25),
			editor.SetCellValue(2, 0, "Q2"),
			editor.SetCellValue(2, 1, 150),
			editor.SetCellValue(2, 2, 40),
			editor.SetCellValue(3, 0, "Q3"),
			editor.SetCellValue(3, 1, 120),
			editor.SetCellValue(3, 2, 30),
			editor.SetCellValue(4, 0, "Q4"),
			editor.SetCellValue(4, 1, 180),
			editor.SetCellValue(4, 2, 50),
			// Add a column chart
			editor.AddChart(editor.ChartTypeColumn, "A1:C5", false, 0, 5, 20, 12,
				editor.WithChartTitle("Quarterly Performance"),
				editor.WithChartStyle(10),
				editor.WithChartLegend(true),
			),
		),
	)
	if err != nil {
		log.Fatalf("edit spreadsheet: %v", err)
	}

	fmt.Println("Created workbook with quarterly data and chart")
	fmt.Println()

	dataSource := datasource.BytesSource(workbookData)

	// Export as PNG
	pngPath := filepath.Join(outDir, "chart.png")
	err = converter.ExportChartToFile(
		dataSource,
		pngPath,
		0, // sheet index
		0, // chart index
		&converter.ChartExportOptions{
			Format: converter.ChartExportFormatPNG,
		},
	)
	if err != nil {
		log.Fatalf("export chart as PNG: %v", err)
	}
	fmt.Printf("✓ Exported chart as PNG: %s\n", pngPath)

	// Export as JPEG with custom quality
	jpegPath := filepath.Join(outDir, "chart.jpg")
	err = converter.ExportChartToFile(
		dataSource,
		jpegPath,
		0, // sheet index
		0, // chart index
		&converter.ChartExportOptions{
			Format:  converter.ChartExportFormatJPEG,
			Quality: 90,
		},
	)
	if err != nil {
		log.Fatalf("export chart as JPEG: %v", err)
	}
	fmt.Printf("✓ Exported chart as JPEG: %s\n", jpegPath)

	// Export as SVG (vector format)
	svgPath := filepath.Join(outDir, "chart.svg")
	err = converter.ExportChartToFile(
		dataSource,
		svgPath,
		0, // sheet index
		0, // chart index
		&converter.ChartExportOptions{
			Format: converter.ChartExportFormatSVG,
		},
	)
	if err != nil {
		log.Fatalf("export chart as SVG: %v", err)
	}
	fmt.Printf("✓ Exported chart as SVG: %s\n", svgPath)

	// Export as a real PDF document
	pdfPath := filepath.Join(outDir, "chart.pdf")
	err = converter.ExportChartToFile(
		dataSource,
		pdfPath,
		0, // sheet index
		0, // chart index
		&converter.ChartExportOptions{
			Format: converter.ChartExportFormatPDF,
		},
	)
	if err != nil {
		log.Fatalf("export chart as PDF: %v", err)
	}
	fmt.Printf("✓ Exported chart as PDF: %s\n", pdfPath)

	// Export at an exact pixel size (both dimensions required)
	customPath := filepath.Join(outDir, "chart-custom-size.png")
	err = converter.ExportChartToFile(
		dataSource,
		customPath,
		0, // sheet index
		0, // chart index
		&converter.ChartExportOptions{
			Format: converter.ChartExportFormatPNG,
			Width:  1200,
			Height: 800,
		},
	)
	if err != nil {
		log.Fatalf("export chart with custom size: %v", err)
	}
	fmt.Printf("✓ Exported chart with custom dimensions: %s\n", customPath)

	// Export to bytes (in-memory)
	pngData, err := converter.ExportChartToBytes(
		dataSource,
		0, // sheet index
		0, // chart index
		&converter.ChartExportOptions{
			Format: converter.ChartExportFormatPNG,
		},
	)
	if err != nil {
		log.Fatalf("export chart to bytes: %v", err)
	}
	fmt.Printf("✓ Exported chart to memory: %d bytes\n", len(pngData))

	fmt.Println()
	fmt.Println("All chart exports completed successfully!")
	fmt.Printf("Output files are in: %s\n", outDir)
}

package tests

import (
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
)

// BenchmarkChartCreation measures the performance of creating a single chart.
func BenchmarkChartCreation(b *testing.B) {
	seed, err := datasource.NewEmptyWorkbook()
	if err != nil {
		b.Fatalf("NewEmptyWorkbook: %v", err)
	}

	// Prepare data
	withData, err := editor.EditSpreadsheet(
		seed,
		editor.WithAddWorksheet("Data"),
		editor.InWorksheet("Data",
			editor.SetCellValue(0, 0, "Category"),
			editor.SetCellValue(0, 1, "Value"),
			editor.SetCellValue(1, 0, "A"),
			editor.SetCellValue(1, 1, 10),
			editor.SetCellValue(2, 0, "B"),
			editor.SetCellValue(2, 1, 20),
			editor.SetCellValue(3, 0, "C"),
			editor.SetCellValue(3, 1, 30),
		),
	)
	if err != nil {
		b.Fatalf("Prepare data: %v", err)
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := editor.EditSpreadsheet(
			datasource.BytesSource(withData),
			editor.InWorksheet("Data",
				editor.AddChart(editor.ChartTypeColumn, "A1:B4", true, 5, 0, 20, 7,
					editor.WithChartTitle("Benchmark"),
				),
			),
		)
		if err != nil {
			b.Fatalf("AddChart: %v", err)
		}
	}
}

// BenchmarkChartWithPreset measures the performance of creating a chart with a preset.
func BenchmarkChartWithPreset(b *testing.B) {
	seed, err := datasource.NewEmptyWorkbook()
	if err != nil {
		b.Fatalf("NewEmptyWorkbook: %v", err)
	}

	// Prepare data
	withData, err := editor.EditSpreadsheet(
		seed,
		editor.WithAddWorksheet("Data"),
		editor.InWorksheet("Data",
			editor.SetCellValue(0, 0, "Category"),
			editor.SetCellValue(0, 1, "Value"),
			editor.SetCellValue(1, 0, "A"),
			editor.SetCellValue(1, 1, 10),
			editor.SetCellValue(2, 0, "B"),
			editor.SetCellValue(2, 1, 20),
			editor.SetCellValue(3, 0, "C"),
			editor.SetCellValue(3, 1, 30),
		),
	)
	if err != nil {
		b.Fatalf("Prepare data: %v", err)
	}

	preset := editor.ProfessionalColumn

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := editor.EditSpreadsheet(
			datasource.BytesSource(withData),
			editor.InWorksheet("Data",
				editor.AddChart(editor.ChartTypeColumn, "A1:B4", true, 5, 0, 20, 7,
					editor.WithChartPreset(preset),
				),
			),
		)
		if err != nil {
			b.Fatalf("AddChart with preset: %v", err)
		}
	}
}

// BenchmarkMultipleCharts measures the performance of creating multiple charts
// in a single workbook.
func BenchmarkMultipleCharts(b *testing.B) {
	seed, err := datasource.NewEmptyWorkbook()
	if err != nil {
		b.Fatalf("NewEmptyWorkbook: %v", err)
	}

	// Prepare data
	withData, err := editor.EditSpreadsheet(
		seed,
		editor.WithAddWorksheet("Data"),
		editor.InWorksheet("Data",
			editor.SetCellValue(0, 0, "Category"),
			editor.SetCellValue(0, 1, "Value"),
			editor.SetCellValue(1, 0, "A"),
			editor.SetCellValue(1, 1, 10),
			editor.SetCellValue(2, 0, "B"),
			editor.SetCellValue(2, 1, 20),
			editor.SetCellValue(3, 0, "C"),
			editor.SetCellValue(3, 1, 30),
		),
	)
	if err != nil {
		b.Fatalf("Prepare data: %v", err)
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := editor.EditSpreadsheet(
			datasource.BytesSource(withData),
			editor.InWorksheet("Data",
				editor.AddChart(editor.ChartTypeColumn, "A1:B4", true, 5, 0, 20, 7),
				editor.AddChart(editor.ChartTypePie, "A1:B4", true, 5, 8, 20, 15),
				editor.AddChart(editor.ChartTypeLine, "A1:B4", true, 22, 0, 37, 7),
			),
		)
		if err != nil {
			b.Fatalf("Add multiple charts: %v", err)
		}
	}
}

// BenchmarkChartModification measures the performance of modifying an existing
// chart.
func BenchmarkChartModification(b *testing.B) {
	seed, err := datasource.NewEmptyWorkbook()
	if err != nil {
		b.Fatalf("NewEmptyWorkbook: %v", err)
	}

	// Create a chart first
	withChart, err := editor.EditSpreadsheet(
		seed,
		editor.WithAddWorksheet("Data"),
		editor.InWorksheet("Data",
			editor.SetCellValue(0, 0, "Category"),
			editor.SetCellValue(0, 1, "Value"),
			editor.SetCellValue(1, 0, "A"),
			editor.SetCellValue(1, 1, 10),
			editor.SetCellValue(2, 0, "B"),
			editor.SetCellValue(2, 1, 20),
			editor.SetCellValue(3, 0, "C"),
			editor.SetCellValue(3, 1, 30),
			editor.AddChart(editor.ChartTypeColumn, "A1:B4", true, 5, 0, 20, 7),
		),
	)
	if err != nil {
		b.Fatalf("Create chart: %v", err)
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := editor.EditSpreadsheet(
			datasource.BytesSource(withChart),
			editor.InWorksheet("Data",
				editor.InChart(0,
					editor.WithChartTitle("Modified"),
					editor.WithChartStyle(7),
				),
			),
		)
		if err != nil {
			b.Fatalf("Modify chart: %v", err)
		}
	}
}

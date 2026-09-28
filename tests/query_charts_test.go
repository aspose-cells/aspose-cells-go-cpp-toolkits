package tests

import (
	"errors"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
)

// TestQueryChartInfo verifies that query.ChartInfo reads back the metadata of a
// chart created with editor.AddChart: type, title, style, legend, bounds, and
// series count.
func TestQueryChartInfo(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 5, 0, 20, 7,
				editor.WithChartTitle("Quarterly sales"),
				editor.WithChartStyle(7),
				editor.WithChartLegend(true),
				editor.WithChartLegendPosition(editor.ChartLegendBottom),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	info, err := query.ChartInfo(datasource.BytesSource(out), 0)
	if err != nil {
		t.Fatalf("ChartInfo: %v", err)
	}

	if info.Index != 0 {
		t.Errorf("Index = %d, want 0", info.Index)
	}
	if info.Type != "column" {
		t.Errorf("Type = %q, want %q", info.Type, "column")
	}
	if info.Title != "Quarterly sales" {
		t.Errorf("Title = %q, want %q", info.Title, "Quarterly sales")
	}
	if !info.TitleVisible {
		t.Error("TitleVisible = false, want true")
	}
	if info.Style != 7 {
		t.Errorf("Style = %d, want 7", info.Style)
	}
	if !info.ShowLegend {
		t.Error("ShowLegend = false, want true")
	}
	if info.LegendPosition != "bottom" {
		t.Errorf("LegendPosition = %q, want %q", info.LegendPosition, "bottom")
	}
	if info.Bounds.TopRow != 5 || info.Bounds.LeftColumn != 0 ||
		info.Bounds.BottomRow != 20 || info.Bounds.RightColumn != 7 {
		t.Errorf("Bounds = (%d,%d)/(%d,%d), want (5,0)/(20,7)",
			info.Bounds.TopRow, info.Bounds.LeftColumn,
			info.Bounds.BottomRow, info.Bounds.RightColumn)
	}
	if info.SeriesCount != 1 {
		t.Errorf("SeriesCount = %d, want 1", info.SeriesCount)
	}
}

// TestQueryChartInfoNotFound verifies that ChartInfo returns ErrChartNotFound
// when the chart index is out of range.
func TestQueryChartInfoNotFound(t *testing.T) {
	// Create a workbook with no charts
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	_, err = query.ChartInfo(datasource.BytesSource(out), 0)
	if err == nil {
		t.Fatal("ChartInfo(0) on a chartless sheet: got nil, want ErrChartNotFound")
	}
	if !errors.Is(err, toolkiterrors.ErrChartNotFound) {
		t.Errorf("ChartInfo(0) error = %v, want ErrChartNotFound", err)
	}
}

// TestQueryChartSeriesData verifies that query.ChartSeriesData reads back the
// series data of a chart: each series' values range, category data, and data
// point count.
func TestQueryChartSeriesData(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:C5", true, 0, 3, 15, 10),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	series, err := query.ChartSeriesData(datasource.BytesSource(out), 0)
	if err != nil {
		t.Fatalf("ChartSeriesData: %v", err)
	}

	// The fixture has data in A1:C5 with column A as text (categories). With
	// byColumn=true, the chart should have 2 series (columns B and C).
	if len(series) != 2 {
		t.Fatalf("series count = %d, want 2", len(series))
	}

	// Verify the first series
	s0 := series[0]
	if s0.Index != 0 {
		t.Errorf("series[0].Index = %d, want 0", s0.Index)
	}
	if s0.Values == "" {
		t.Error("series[0].Values is empty, want a cell range")
	}
	if s0.CategoryData == "" {
		t.Error("series[0].CategoryData is empty, want a cell range")
	}
	if s0.DataValueCount != 4 {
		t.Errorf("series[0].DataValueCount = %d, want 4", s0.DataValueCount)
	}

	// Verify the second series
	s1 := series[1]
	if s1.Index != 1 {
		t.Errorf("series[1].Index = %d, want 1", s1.Index)
	}
	if s1.Values == "" {
		t.Error("series[1].Values is empty, want a cell range")
	}
	// Both series should share the same category data
	if s1.CategoryData != s0.CategoryData {
		t.Errorf("series[1].CategoryData = %q, want %q (same as series[0])",
			s1.CategoryData, s0.CategoryData)
	}
}

// TestQueryChartCount verifies that query.ChartCount returns the number of
// charts on the worksheet.
func TestQueryChartCount(t *testing.T) {
	// Create a worksheet with 2 charts
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10),
			editor.AddChart(editor.ChartTypePie, "A1:B5", true, 20, 3, 35, 10),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	count, err := query.ChartCount(datasource.BytesSource(out))
	if err != nil {
		t.Fatalf("ChartCount: %v", err)
	}
	if count != 2 {
		t.Errorf("ChartCount = %d, want 2", count)
	}
}

// TestQueryChartCountEmpty verifies that ChartCount returns 0 for a worksheet
// with no charts.
func TestQueryChartCountEmpty(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	count, err := query.ChartCount(datasource.BytesSource(out))
	if err != nil {
		t.Fatalf("ChartCount: %v", err)
	}
	if count != 0 {
		t.Errorf("ChartCount = %d, want 0", count)
	}
}

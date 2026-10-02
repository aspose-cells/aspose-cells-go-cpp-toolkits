package editor

import (
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// ChartType names a chart type. The constants below are the curated,
// documented spellings; any of the engine's 81 chart types is also accepted by
// its own name (for example ChartType("pyramid") or ChartType("radarFilled")),
// and names are matched case- and punctuation-insensitively, so "Column3D",
// "column_3d" and "COLUMN 3D" are the same type. An unrecognized name yields
// ErrInvalidChartType rather than falling back to a default.
type ChartType string

// The curated chart types. Each names a family; the stacked and 100%-stacked
// variants follow the same "…Stacked" / "…100PercentStacked" pattern, and every
// remaining engine type is reachable by raw name.
const (
	ChartTypeColumn           ChartType = "column"
	ChartTypeColumnStacked    ChartType = "columnStacked"
	ChartTypeColumn3D         ChartType = "column3D"
	ChartTypeBar              ChartType = "bar"
	ChartTypeBarStacked       ChartType = "barStacked"
	ChartTypeBar3DClustered   ChartType = "bar3DClustered"
	ChartTypeLine             ChartType = "line"
	ChartTypeLineStacked      ChartType = "lineStacked"
	ChartTypeLineWithMarkers  ChartType = "lineWithDataMarkers"
	ChartTypeLine3D           ChartType = "line3D"
	ChartTypeArea             ChartType = "area"
	ChartTypeAreaStacked      ChartType = "areaStacked"
	ChartTypeArea3D           ChartType = "area3D"
	ChartTypePie              ChartType = "pie"
	ChartTypePie3D            ChartType = "pie3D"
	ChartTypePieExploded      ChartType = "pieExploded"
	ChartTypePieBar           ChartType = "pieBar"
	ChartTypeDoughnut         ChartType = "doughnut"
	ChartTypeDoughnutExploded ChartType = "doughnutExploded"
	ChartTypeScatter          ChartType = "scatter"
	ChartTypeBubble           ChartType = "bubble"
	ChartTypeRadar            ChartType = "radar"
	ChartTypeRadarFilled      ChartType = "radarFilled"
	ChartTypeStock            ChartType = "stockHighLowClose"
	ChartTypeSurface3D        ChartType = "surface3D"
	ChartTypeSurfaceContour   ChartType = "surfaceContour"
	ChartTypeBoxWhisker       ChartType = "boxWhisker"
	ChartTypeFunnel           ChartType = "funnel"
	ChartTypeParetoLine       ChartType = "paretoLine"
	ChartTypeSunburst         ChartType = "sunburst"
	ChartTypeTreemap          ChartType = "treemap"
	ChartTypeWaterfall        ChartType = "waterfall"
	ChartTypeHistogram        ChartType = "histogram"
	ChartTypeMap              ChartType = "map"
)

// ChartLegendPosition names where a chart's legend is docked. Like ChartType it
// is matched case- and punctuation-insensitively.
type ChartLegendPosition string

// The legend positions. The default for a newly added chart is
// ChartLegendRight.
//
// ChartLegendNotDocked is the one exception to "what you set is what the file
// holds": the engine accepts it and reports it back while the workbook is in
// memory, but XLSX has no way to record an undocked legend, so a chart saved with
// it reloads docked right like any chart that never had a position set. Use it
// only when the workbook is not going to be saved and reloaded.
const (
	ChartLegendBottom    ChartLegendPosition = "bottom"
	ChartLegendCorner    ChartLegendPosition = "corner"
	ChartLegendTop       ChartLegendPosition = "top"
	ChartLegendRight     ChartLegendPosition = "right"
	ChartLegendLeft      ChartLegendPosition = "left"
	ChartLegendNotDocked ChartLegendPosition = "notDocked"
)

// InChart creates a WorksheetAction that applies chart actions to the worksheet's
// existing chart at the given index. Charts are indexed from zero in the order
// they were added.
//
// Parameters:
//   - index: The zero-based index of the chart on the worksheet.
//   - actions: A variadic list of ChartAction functions to apply to that chart.
//
// Returns:
//   - WorksheetAction: A function that modifies the chart. Returns
//     ErrChartNotFound when no chart has that index.
//
// Example:
//
//	editor.InWorksheet(0, editor.InChart(0,
//		editor.WithChartTitle("Revised"),
//		editor.WithChartStyle(12),
//	))
func InChart(index int, actions ...ChartAction) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		chart, err := cells.Chart(worksheet, index)
		if err != nil {
			return err
		}
		for _, action := range actions {
			if err := action(chart); err != nil {
				return err
			}
		}
		return nil
	}
}

// WithChartBounds creates a ChartAction that positions the chart over a cell
// rectangle.
//
// Parameters:
//   - topRow: The zero-based row of the chart's top edge.
//   - leftColumn: The zero-based column of the chart's left edge.
//   - bottomRow: The zero-based row of the chart's bottom edge.
//   - rightColumn: The zero-based column of the chart's right edge.
//
// Returns:
//   - ChartAction: A function that repositions the chart. Returns ErrInvalidRange
//     when the rectangle is reversed, negative, or outside the worksheet grid.
func WithChartBounds(topRow, leftColumn, bottomRow, rightColumn int) ChartAction {
	return func(chart *asposecells.Chart) error {
		return cells.MoveChart(chart, topRow, leftColumn, bottomRow, rightColumn)
	}
}

// WithChartType creates a ChartAction that switches the chart to a different
// type. Changing the type reinterprets the existing data, so a pie chart and a
// column chart show the same series differently rather than needing new data.
//
// Parameters:
//   - chartType: The new chart type. See ChartType for the accepted names.
//
// Returns:
//   - ChartAction: A function that changes the chart type. Returns
//     ErrInvalidChartType for an unknown name.
func WithChartType(chartType ChartType) ChartAction {
	return func(chart *asposecells.Chart) error {
		resolved, err := cells.ResolveChartType(string(chartType))
		if err != nil {
			return err
		}
		return chart.SetType(resolved)
	}
}

// WithChartTitle creates a ChartAction that sets the chart's title text and makes
// the title visible. Visibility is set along with the text because it does not
// survive a save and reload on its own: the engine writes a title that reads
// back as hidden the next time the file is loaded.
//
// Parameters:
//   - text: The title text.
//
// Returns:
//   - ChartAction: A function that sets the chart title.
func WithChartTitle(text string) ChartAction {
	return func(chart *asposecells.Chart) error {
		return cells.SetChartTitle(chart, text)
	}
}

// HideChartTitle creates a ChartAction that hides the chart's title.
//
// Hiding is not reversible through this API: the engine also discards the title's
// text when it is hidden, so a chart that is hidden and later reloaded reports the
// automatic title derived from its series rather than the text that was set
// before. Use WithChartTitle to (re)set a title, including on a chart whose title
// is currently hidden.
//
// Returns:
//   - ChartAction: A function that hides the chart title.
func HideChartTitle() ChartAction {
	return func(chart *asposecells.Chart) error {
		return cells.HideChartTitle(chart)
	}
}

// WithChartStyle creates a ChartAction that applies one of the engine's built-in
// chart styles, which colour and format the chart as a whole.
//
// Parameters:
//   - style: The style number, from 1 to 48.
//
// Returns:
//   - ChartAction: A function that applies the style. Returns
//     ErrInvalidChartStyle when style is outside 1..48.
func WithChartStyle(style int) ChartAction {
	return func(chart *asposecells.Chart) error {
		return cells.SetChartStyle(chart, style)
	}
}

// WithChartLegend creates a ChartAction that shows or hides the chart's legend.
//
// Parameters:
//   - visible: Whether the legend is shown.
//
// Returns:
//   - ChartAction: A function that sets the legend's visibility.
func WithChartLegend(visible bool) ChartAction {
	return func(chart *asposecells.Chart) error {
		return cells.SetChartLegend(chart, visible)
	}
}

// WithChartLegendPosition creates a ChartAction that docks the chart's legend at
// one of the engine's positions. It does not by itself show a hidden legend;
// pair it with WithChartLegend(true) when the legend starts hidden.
//
// Measured, every position but ChartLegendNotDocked survives a save and reload
// unchanged. ChartLegendNotDocked is applied and reads back correctly in memory,
// but an XLSX file cannot record an undocked legend, so a reloaded chart reports
// the default position instead. See ChartLegendPosition.
//
// Parameters:
//   - position: The legend position. See ChartLegendPosition.
//
// Returns:
//   - ChartAction: A function that moves the legend. Returns
//     ErrInvalidChartPosition for an unknown position name.
func WithChartLegendPosition(position ChartLegendPosition) ChartAction {
	return func(chart *asposecells.Chart) error {
		return cells.SetChartLegendPosition(chart, string(position))
	}
}

// WithChartDataRange creates a ChartAction that re-points the chart at a
// different source range. This is the way to change a chart's data wholesale:
// the series are rebuilt from the new range.
//
// Parameters:
//   - dataRange: The new source range, e.g. "A1:C5".
//   - byColumn: Whether each column of dataRange is a series (true) or each row is.
//
// Returns:
//   - ChartAction: A function that changes the chart's source data. Returns
//     ErrInvalidRange when dataRange is not a plottable in-grid range.
func WithChartDataRange(dataRange string, byColumn bool) ChartAction {
	return func(chart *asposecells.Chart) error {
		return cells.SetChartDataRange(chart, dataRange, byColumn)
	}
}

// WithChartSeries creates a ChartAction that appends one more series to the
// chart, in addition to the ones the chart already has. Use it to build a chart
// from several separate ranges; use WithChartDataRange to replace them all.
//
// Parameters:
//   - dataRange: The series' source range, e.g. "B2:B5".
//   - byColumn: Whether the series runs down the range's column (true) or along
//     its row (false).
//
// Returns:
//   - ChartAction: A function that adds the series. Returns ErrInvalidRange when
//     dataRange is not a plottable in-grid range.
func WithChartSeries(dataRange string, byColumn bool) ChartAction {
	return func(chart *asposecells.Chart) error {
		return cells.AddChartSeries(chart, dataRange, byColumn)
	}
}

// RemoveChartSeries creates a ChartAction that removes the chart's series at the
// given index. The series after it shift down.
//
// Parameters:
//   - index: The zero-based index of the series to remove.
//
// Returns:
//   - ChartAction: A function that removes the series. Returns ErrChartNotFound
//     when the chart has no series with that index.
func RemoveChartSeries(index int) ChartAction {
	return func(chart *asposecells.Chart) error {
		return cells.RemoveChartSeriesAt(chart, index)
	}
}

// ClearChartSeries creates a ChartAction that removes every series from the
// chart, leaving it with its formatting and title but no data.
//
// Returns:
//   - ChartAction: A function that clears the chart's series.
func ClearChartSeries() ChartAction {
	return func(chart *asposecells.Chart) error {
		return cells.ClearChartSeries(chart)
	}
}

// WithChartCategoryData creates a ChartAction that sets the range the chart's
// category axis labels come from. Category data set this way survives a later
// WithChartDataRange rebuild of the series.
//
// Parameters:
//   - dataRange: The source range for the category labels, e.g. "A2:A5".
//
// Returns:
//   - ChartAction: A function that sets the category data. Returns ErrInvalidRange
//     when dataRange is not a plottable in-grid range.
func WithChartCategoryData(dataRange string) ChartAction {
	return func(chart *asposecells.Chart) error {
		return cells.SetChartCategoryData(chart, dataRange)
	}
}

// ChartStylePreset bundles common chart styling and configuration options into
// a single value for reuse across multiple charts or workbooks. Rather than
// applying WithChartType, WithChartStyle, WithChartTitle, WithChartLegend,
// WithChartLegendPosition, WithChartDataRange, WithChartCategoryData, and
// WithChartBounds separately, a preset combines them.
//
// A preset can be partial: zero-value fields are skipped when applied, so a
// preset can specify only the settings it cares about and leave the rest
// unchanged. This makes presets useful for both complete chart templates and
// targeted style bundles.
type ChartStylePreset struct {
	// ChartType is the chart's type (e.g., ChartTypeColumn, ChartTypePie).
	// Empty means the chart type is not changed.
	ChartType ChartType

	// Style is the built-in chart style number (1..48). Zero means no style is
	// applied, leaving the chart at its default or previously set style.
	Style int

	// Title is the chart's title text. Empty means no title is set; the chart
	// may still show an automatic title derived from the series.
	Title string

	// ShowLegend controls whether the legend is visible.
	ShowLegend bool

	// LegendPosition is where the legend is docked. Empty means the legend
	// position is not changed; it is only applied when ShowLegend is true.
	LegendPosition ChartLegendPosition

	// DataRange is the chart's source data range in A1 notation (e.g., "A1:C5").
	// Empty means the data range is not changed.
	DataRange string

	// ByColumn selects how the data range is read into series. True reads down
	// the columns (one series per column), false reads across the rows. Only
	// applied when DataRange is non-empty.
	ByColumn bool

	// CategoryData is the cell range for the category axis labels (e.g.,
	// "A2:A5"). Empty means the category data is not explicitly set; the engine
	// derives it from the data range. Only applied when non-empty.
	CategoryData string

	// Bounds is the chart's position on the worksheet: (topRow, leftColumn,
	// bottomRow, rightColumn), all zero-based. A zero-value Bounds (all fields
	// zero) means the position is not changed. Only applied when at least one
	// field is non-zero.
	Bounds ChartBounds
}

// ChartBounds is a chart's position on the worksheet, in zero-based row and
// column coordinates. Used by ChartStylePreset to specify where the chart is
// placed.
type ChartBounds struct {
	TopRow      int
	LeftColumn  int
	BottomRow   int
	RightColumn int
}

// Predefined chart templates for common use cases. Each template is a
// ChartStylePreset with sensible defaults for a particular chart style. Use
// them directly or as a starting point for custom presets.
//
// Example:
//
//	editor.InChart(0, editor.WithChartPreset(editor.ProfessionalColumn))
var (
	// ProfessionalColumn is a clean, professional column chart with a title,
	// bottom legend, and built-in style 7. Suitable for business reports.
	ProfessionalColumn = ChartStylePreset{
		ChartType:      ChartTypeColumn,
		Style:          7,
		ShowLegend:     true,
		LegendPosition: ChartLegendBottom,
	}

	// MinimalPie is a simple pie chart without a legend, relying on data
	// labels. Suitable for presentations where space is limited.
	MinimalPie = ChartStylePreset{
		ChartType:  ChartTypePie,
		Style:      3,
		ShowLegend: false,
	}

	// PresentationBar is a bar chart optimized for presentations: clear title,
	// right legend, and built-in style 10.
	PresentationBar = ChartStylePreset{
		ChartType:      ChartTypeBar,
		Style:          10,
		ShowLegend:     true,
		LegendPosition: ChartLegendRight,
	}

	// DashboardLine is a line chart with markers, suitable for dashboards.
	// Style 12 provides good visibility on screens.
	DashboardLine = ChartStylePreset{
		ChartType:      ChartTypeLineWithMarkers,
		Style:          12,
		ShowLegend:     true,
		LegendPosition: ChartLegendTop,
	}

	// ReportArea is an area chart for showing trends over time in reports.
	// Style 5 provides a clean look.
	ReportArea = ChartStylePreset{
		ChartType:      ChartTypeArea,
		Style:          5,
		ShowLegend:     true,
		LegendPosition: ChartLegendBottom,
	}

	// SimpleScatter is a scatter plot without a legend, suitable for showing
	// correlations.
	SimpleScatter = ChartStylePreset{
		ChartType:  ChartTypeScatter,
		Style:      8,
		ShowLegend: false,
	}

	// FinancialCandlestick is a candlestick chart for financial data, showing
	// open/high/low/close values. Style 20 provides good visibility.
	FinancialCandlestick = ChartStylePreset{
		ChartType:  ChartTypeStock,
		Style:      20,
		ShowLegend: false,
	}

	// ComparisonColumn3D is a 3D column chart for comparing multiple series.
	// Style 15 provides depth perception.
	ComparisonColumn3D = ChartStylePreset{
		ChartType:      ChartTypeColumn3D,
		Style:          15,
		ShowLegend:     true,
		LegendPosition: ChartLegendRight,
	}

	// TrendLine is a line chart optimized for showing trends over time. Style
	// 12 with markers highlights data points.
	TrendLine = ChartStylePreset{
		ChartType:      ChartTypeLineWithMarkers,
		Style:          12,
		ShowLegend:     true,
		LegendPosition: ChartLegendBottom,
	}

	// DistributionDoughnut is a doughnut chart for showing distributions. Style
	// 6 provides clear segment separation.
	DistributionDoughnut = ChartStylePreset{
		ChartType:      ChartTypeDoughnut,
		Style:          6,
		ShowLegend:     true,
		LegendPosition: ChartLegendRight,
	}

	// StackedArea is a stacked area chart for showing part-to-whole
	// relationships over time. Style 5 provides clean stacking.
	StackedArea = ChartStylePreset{
		ChartType:      ChartTypeAreaStacked,
		Style:          5,
		ShowLegend:     true,
		LegendPosition: ChartLegendBottom,
	}

	// RadarComparison is a radar chart for comparing multiple variables across
	// categories. Style 18 provides good visibility.
	RadarComparison = ChartStylePreset{
		ChartType:      ChartTypeRadar,
		Style:          18,
		ShowLegend:     true,
		LegendPosition: ChartLegendTop,
	}

	// Industry-specific templates

	// SalesPerformance is a column chart optimized for sales reports. Shows
	// multiple product lines or regions with clear comparison.
	SalesPerformance = ChartStylePreset{
		ChartType:      ChartTypeColumn,
		Style:          7,
		ShowLegend:     true,
		LegendPosition: ChartLegendBottom,
	}

	// FinancialSummary is a professional bar chart for financial statements.
	// Clean layout suitable for annual reports.
	FinancialSummary = ChartStylePreset{
		ChartType:      ChartTypeBar,
		Style:          10,
		ShowLegend:     true,
		LegendPosition: ChartLegendRight,
	}

	// MarketShare is a pie chart for showing market distribution. No legend,
	// relies on data labels for clarity.
	MarketShare = ChartStylePreset{
		ChartType:  ChartTypePie,
		Style:      3,
		ShowLegend: false,
	}

	// ScientificData is a scatter plot for scientific measurements. Clean,
	// minimal style with markers.
	ScientificData = ChartStylePreset{
		ChartType:  ChartTypeScatter,
		Style:      8,
		ShowLegend: false,
	}

	// ProjectTimeline is a line chart for project milestones and timelines.
	// Markers highlight key dates.
	ProjectTimeline = ChartStylePreset{
		ChartType:      ChartTypeLineWithMarkers,
		Style:          12,
		ShowLegend:     true,
		LegendPosition: ChartLegendTop,
	}

	// SurveyResults is a horizontal bar chart for survey responses. Easy to
	// read category labels.
	SurveyResults = ChartStylePreset{
		ChartType:  ChartTypeBar,
		Style:      9,
		ShowLegend: false,
	}

	// BudgetVariance is a column chart for budget vs actual comparisons.
	// Clear visual distinction between planned and actual values.
	BudgetVariance = ChartStylePreset{
		ChartType:      ChartTypeColumn,
		Style:          11,
		ShowLegend:     true,
		LegendPosition: ChartLegendBottom,
	}

	// KPIDashboard is a compact line chart for key performance indicators.
	// Optimized for dashboard displays.
	KPIDashboard = ChartStylePreset{
		ChartType:  ChartTypeLine,
		Style:      12,
		ShowLegend: false,
	}
)

// WithChartPreset creates a ChartAction that applies a ChartStylePreset's
// configuration and styling options to the chart. The preset's fields are
// applied in order: chart type, data range, category data, bounds, style,
// title, legend visibility, and legend position. A zero-value field
// (ChartType="", Style=0, Title="", DataRange="", CategoryData="", Bounds with
// all fields zero, LegendPosition="") is skipped, so a preset can partially
// specify configuration and leave the rest unchanged.
//
// Parameters:
//   - preset: The preset to apply.
//
// Returns:
//   - ChartAction: A function that applies the preset. Returns
//     ErrInvalidChartType for an unknown chart type name, ErrInvalidChartStyle
//     when preset.Style is outside 1..48 (unless zero), ErrInvalidRange for
//     malformed data/category ranges or reversed bounds, and
//     ErrInvalidChartPosition for an unknown legend position name.
//
// Example:
//
//	preset := editor.ChartStylePreset{
//	    ChartType:      editor.ChartTypeColumn,
//	    Style:          7,
//	    Title:          "Quarterly sales",
//	    ShowLegend:     true,
//	    LegendPosition: editor.ChartLegendBottom,
//	    DataRange:      "A1:C5",
//	    ByColumn:       true,
//	    CategoryData:   "A2:A5",
//	    Bounds:         editor.ChartBounds{TopRow: 5, LeftColumn: 0, BottomRow: 20, RightColumn: 7},
//	}
//	editor.InChart(0, editor.WithChartPreset(preset))
func WithChartPreset(preset ChartStylePreset) ChartAction {
	return func(chart *asposecells.Chart) error {
		// Chart type
		if preset.ChartType != "" {
			resolved, err := cells.ResolveChartType(string(preset.ChartType))
			if err != nil {
				return err
			}
			if err := chart.SetType(resolved); err != nil {
				return err
			}
		}

		// Data range
		if preset.DataRange != "" {
			if err := cells.SetChartDataRange(chart, preset.DataRange, preset.ByColumn); err != nil {
				return err
			}
		}

		// Category data
		if preset.CategoryData != "" {
			if err := cells.SetChartCategoryData(chart, preset.CategoryData); err != nil {
				return err
			}
		}

		// Bounds
		if preset.Bounds.TopRow != 0 || preset.Bounds.LeftColumn != 0 ||
			preset.Bounds.BottomRow != 0 || preset.Bounds.RightColumn != 0 {
			if err := cells.MoveChart(chart, preset.Bounds.TopRow, preset.Bounds.LeftColumn,
				preset.Bounds.BottomRow, preset.Bounds.RightColumn); err != nil {
				return err
			}
		}

		// Style
		if preset.Style != 0 {
			if err := cells.SetChartStyle(chart, preset.Style); err != nil {
				return err
			}
		}

		// Title
		if preset.Title != "" {
			if err := cells.SetChartTitle(chart, preset.Title); err != nil {
				return err
			}
		}

		// Legend
		if err := cells.SetChartLegend(chart, preset.ShowLegend); err != nil {
			return err
		}
		if preset.ShowLegend && preset.LegendPosition != "" {
			if err := cells.SetChartLegendPosition(chart, string(preset.LegendPosition)); err != nil {
				return err
			}
		}
		return nil
	}
}

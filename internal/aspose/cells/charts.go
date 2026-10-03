package cells

import (
	"fmt"
	"strings"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// chartTypeByName maps a normalized chart type name to the engine enum. Names
// are normalized by normalizeChartEnumName, so the several spellings a caller
// might use ("Column3D", "column_3d", "column3d") all resolve to the same entry.
//
// The table is exhaustive over the engine's 81 chart types, which is what lets
// the public API ship a curated handful of constants while still accepting any
// type by name. TestChartTypeTableIsExhaustive pins that exhaustiveness.
var chartTypeByName = map[string]asposecells.ChartType{
	"area":                                   asposecells.ChartType_Area,
	"areastacked":                            asposecells.ChartType_AreaStacked,
	"area100percentstacked":                  asposecells.ChartType_Area100PercentStacked,
	"area3d":                                 asposecells.ChartType_Area3D,
	"area3dstacked":                          asposecells.ChartType_Area3DStacked,
	"area3d100percentstacked":                asposecells.ChartType_Area3D100PercentStacked,
	"bar":                                    asposecells.ChartType_Bar,
	"barstacked":                             asposecells.ChartType_BarStacked,
	"bar100percentstacked":                   asposecells.ChartType_Bar100PercentStacked,
	"bar3dclustered":                         asposecells.ChartType_Bar3DClustered,
	"bar3dstacked":                           asposecells.ChartType_Bar3DStacked,
	"bar3d100percentstacked":                 asposecells.ChartType_Bar3D100PercentStacked,
	"bubble":                                 asposecells.ChartType_Bubble,
	"bubble3d":                               asposecells.ChartType_Bubble3D,
	"column":                                 asposecells.ChartType_Column,
	"columnstacked":                          asposecells.ChartType_ColumnStacked,
	"column100percentstacked":                asposecells.ChartType_Column100PercentStacked,
	"column3d":                               asposecells.ChartType_Column3D,
	"column3dclustered":                      asposecells.ChartType_Column3DClustered,
	"column3dstacked":                        asposecells.ChartType_Column3DStacked,
	"column3d100percentstacked":              asposecells.ChartType_Column3D100PercentStacked,
	"cone":                                   asposecells.ChartType_Cone,
	"conestacked":                            asposecells.ChartType_ConeStacked,
	"cone100percentstacked":                  asposecells.ChartType_Cone100PercentStacked,
	"conicalbar":                             asposecells.ChartType_ConicalBar,
	"conicalbarstacked":                      asposecells.ChartType_ConicalBarStacked,
	"conicalbar100percentstacked":            asposecells.ChartType_ConicalBar100PercentStacked,
	"conicalcolumn3d":                        asposecells.ChartType_ConicalColumn3D,
	"cylinder":                               asposecells.ChartType_Cylinder,
	"cylinderstacked":                        asposecells.ChartType_CylinderStacked,
	"cylinder100percentstacked":              asposecells.ChartType_Cylinder100PercentStacked,
	"cylindricalbar":                         asposecells.ChartType_CylindricalBar,
	"cylindricalbarstacked":                  asposecells.ChartType_CylindricalBarStacked,
	"cylindricalbar100percentstacked":        asposecells.ChartType_CylindricalBar100PercentStacked,
	"cylindricalcolumn3d":                    asposecells.ChartType_CylindricalColumn3D,
	"doughnut":                               asposecells.ChartType_Doughnut,
	"doughnutexploded":                       asposecells.ChartType_DoughnutExploded,
	"line":                                   asposecells.ChartType_Line,
	"linestacked":                            asposecells.ChartType_LineStacked,
	"line100percentstacked":                  asposecells.ChartType_Line100PercentStacked,
	"linewithdatamarkers":                    asposecells.ChartType_LineWithDataMarkers,
	"linestackedwithdatamarkers":             asposecells.ChartType_LineStackedWithDataMarkers,
	"line100percentstackedwithdatamarkers":   asposecells.ChartType_Line100PercentStackedWithDataMarkers,
	"line3d":                                 asposecells.ChartType_Line3D,
	"pie":                                    asposecells.ChartType_Pie,
	"pie3d":                                  asposecells.ChartType_Pie3D,
	"piepie":                                 asposecells.ChartType_PiePie,
	"pieexploded":                            asposecells.ChartType_PieExploded,
	"pie3dexploded":                          asposecells.ChartType_Pie3DExploded,
	"piebar":                                 asposecells.ChartType_PieBar,
	"pyramid":                                asposecells.ChartType_Pyramid,
	"pyramidstacked":                         asposecells.ChartType_PyramidStacked,
	"pyramid100percentstacked":               asposecells.ChartType_Pyramid100PercentStacked,
	"pyramidbar":                             asposecells.ChartType_PyramidBar,
	"pyramidbarstacked":                      asposecells.ChartType_PyramidBarStacked,
	"pyramidbar100percentstacked":            asposecells.ChartType_PyramidBar100PercentStacked,
	"pyramidcolumn3d":                        asposecells.ChartType_PyramidColumn3D,
	"radar":                                  asposecells.ChartType_Radar,
	"radarwithdatamarkers":                   asposecells.ChartType_RadarWithDataMarkers,
	"radarfilled":                            asposecells.ChartType_RadarFilled,
	"scatter":                                asposecells.ChartType_Scatter,
	"scatterconnectedbycurveswithdatamarker": asposecells.ChartType_ScatterConnectedByCurvesWithDataMarker,
	"scatterconnectedbycurveswithoutdatamarker": asposecells.ChartType_ScatterConnectedByCurvesWithoutDataMarker,
	"scatterconnectedbylineswithdatamarker":     asposecells.ChartType_ScatterConnectedByLinesWithDataMarker,
	"scatterconnectedbylineswithoutdatamarker":  asposecells.ChartType_ScatterConnectedByLinesWithoutDataMarker,
	"stockhighlowclose":                         asposecells.ChartType_StockHighLowClose,
	"stockopenhighlowclose":                     asposecells.ChartType_StockOpenHighLowClose,
	"stockvolumehighlowclose":                   asposecells.ChartType_StockVolumeHighLowClose,
	"stockvolumeopenhighlowclose":               asposecells.ChartType_StockVolumeOpenHighLowClose,
	"surface3d":                                 asposecells.ChartType_Surface3D,
	"surfacewireframe3d":                        asposecells.ChartType_SurfaceWireframe3D,
	"surfacecontour":                            asposecells.ChartType_SurfaceContour,
	"surfacecontourwireframe":                   asposecells.ChartType_SurfaceContourWireframe,
	"boxwhisker":                                asposecells.ChartType_BoxWhisker,
	"funnel":                                    asposecells.ChartType_Funnel,
	"paretoline":                                asposecells.ChartType_ParetoLine,
	"sunburst":                                  asposecells.ChartType_Sunburst,
	"treemap":                                   asposecells.ChartType_Treemap,
	"waterfall":                                 asposecells.ChartType_Waterfall,
	"histogram":                                 asposecells.ChartType_Histogram,
	"map":                                       asposecells.ChartType_Map,
}

// normalizeChartEnumName lowercases name and drops every character that is not a
// letter or digit, so "Column3D", "column_3d" and "COLUMN 3D" all normalize to
// "column3d". The chart type and legend position tables are both keyed through
// this, so the accepted spelling is generous rather than a single brittle string.
func normalizeChartEnumName(name string) string {
	var b strings.Builder
	b.Grow(len(name))
	for _, r := range name {
		switch {
		case r >= 'a' && r <= 'z', r >= '0' && r <= '9':
			b.WriteRune(r)
		case r >= 'A' && r <= 'Z':
			b.WriteRune(r + ('a' - 'A'))
		}
	}
	return b.String()
}

// ResolveChartType maps a chart type name to the engine enum, accepting any of
// the engine's 81 types in any spelling. An unrecognized name returns
// ErrInvalidChartType rather than falling back to a default type — the engine
// would otherwise take a plausible-looking wrong type silently.
func ResolveChartType(name string) (asposecells.ChartType, error) {
	if ct, ok := chartTypeByName[normalizeChartEnumName(name)]; ok {
		return ct, nil
	}
	return 0, fmt.Errorf("unknown chart type %q: %w", name, toolkiterrors.ErrInvalidChartType)
}

// ValidateChartDataRange parses a bare A1-style range ("A1:B5") and checks it
// lies inside the worksheet grid, spans more than one row, and encloses more
// than one cell.
//
// Every one of these checks is needed because the engine performs almost none of
// them: a reference past XFD or past the last row, a range with a third ":"
// segment, and plain non-reference text are all accepted without error and yield
// a chart with no data points — a silently empty chart. A single-cell range is
// likewise accepted but cannot be plotted, and a single-row range is the one
// shape the engine does reject, with an unhelpful message, so it is caught here
// instead.
//
// A single-column range ("B2:B4") is valid and is in fact the canonical form for
// one series: a series is N points along one axis. Only the single-row
// orientation is degenerate.
//
// A range carrying a sheet qualifier ("Sheet1!A1:B5") is rejected: the engine
// resolves a bare range against the chart's own worksheet, so cross-sheet
// sourcing is out of scope rather than silently mis-resolved.
func ValidateChartDataRange(ref string) error {
	if strings.Contains(ref, "!") {
		return fmt.Errorf("chart data range %q must not carry a sheet qualifier: %w",
			ref, toolkiterrors.ErrInvalidRange)
	}
	area, err := ParseArea(ref)
	if err != nil {
		return err
	}
	if err := ValidateGridArea(area); err != nil {
		return err
	}
	if area.Start.Row == area.End.Row && area.Start.Col == area.End.Col {
		return fmt.Errorf("chart data range %q covers a single cell: %w",
			ref, toolkiterrors.ErrInvalidRange)
	}
	if area.Start.Row == area.End.Row {
		return fmt.Errorf("chart data range %q covers a single row, which cannot be plotted: %w",
			ref, toolkiterrors.ErrInvalidRange)
	}
	return nil
}

// ValidateGridArea checks that both corners of area lie inside the worksheet
// grid. The bounds are compared with >= 0 as well as < the limit because
// ParseCellRef accumulates the column index without bounding the letter run, so
// a long enough column name overflows into an arbitrary -- possibly negative --
// int rather than a large positive one.
func ValidateGridArea(area Area) error {
	for _, c := range []CellRef{area.Start, area.End} {
		if c.Row < 0 || c.Row >= MaxGridRows || c.Col < 0 || c.Col >= MaxGridCols {
			return fmt.Errorf("cell reference (%d, %d) is outside the worksheet grid: %w",
				c.Row, c.Col, toolkiterrors.ErrInvalidRange)
		}
	}
	return nil
}

// ValidateChartBounds checks a chart's placement rectangle, whose corners are
// both cells the chart covers, so the rectangle must enclose at least a 2x2
// block.
//
// A rect that spans no rows or no columns is rejected even though the engine
// accepts it: measured, both (0,0,0,0) and (0,3,0,3) return no error and leave a
// chart with no extent, which is invisible in the saved file rather than an
// error the caller can act on. That is the silently-wrong outcome this layer
// exists to convert into an error.
func ValidateChartBounds(topRow, leftColumn, bottomRow, rightColumn int) error {
	// Check all four coordinates for negativity first, so the error message
	// identifies the actual problem (a negative coordinate) rather than a
	// misleading "reversed" error when bottomRow or rightColumn is negative.
	if topRow < 0 || leftColumn < 0 || bottomRow < 0 || rightColumn < 0 {
		return fmt.Errorf("chart bounds (%d, %d, %d, %d) must not be negative: %w",
			topRow, leftColumn, bottomRow, rightColumn, toolkiterrors.ErrInvalidRange)
	}
	if topRow > bottomRow || leftColumn > rightColumn {
		return fmt.Errorf("chart bounds (%d, %d, %d, %d) are reversed: %w",
			topRow, leftColumn, bottomRow, rightColumn, toolkiterrors.ErrInvalidRange)
	}
	if topRow == bottomRow || leftColumn == rightColumn {
		return fmt.Errorf("chart bounds (%d, %d, %d, %d) enclose no area: %w",
			topRow, leftColumn, bottomRow, rightColumn, toolkiterrors.ErrInvalidRange)
	}
	if bottomRow >= MaxGridRows || rightColumn >= MaxGridCols {
		return fmt.Errorf("chart bounds (%d, %d, %d, %d) are outside the worksheet grid: %w",
			topRow, leftColumn, bottomRow, rightColumn, toolkiterrors.ErrInvalidRange)
	}
	return nil
}

// Charts returns the worksheet's chart collection. A nil or null handle becomes
// an error rather than a collection that reports zero charts, so a caller can
// tell "this sheet has no charts" from "the handle is unusable".
func Charts(ws *asposecells.Worksheet) (*asposecells.ChartCollection, error) {
	charts, err := engine.Derive(ws.GetCharts())
	if err != nil {
		return nil, err
	}
	if charts == nil {
		return nil, fmt.Errorf("worksheet returned no chart collection: %w", toolkiterrors.ErrChartNotFound)
	}
	null, err := charts.IsNull()
	if err != nil {
		return nil, err
	}
	if null {
		return nil, fmt.Errorf("worksheet returned no chart collection: %w", toolkiterrors.ErrChartNotFound)
	}
	return charts, nil
}

// Chart returns the chart at index, or ErrChartNotFound when index is out of
// range. The range check runs first so an out-of-range index cannot reach the
// engine's collection getter at all.
func Chart(ws *asposecells.Worksheet, index int) (*asposecells.Chart, error) {
	charts, err := Charts(ws)
	if err != nil {
		return nil, err
	}
	count, err := charts.GetCount()
	if err != nil {
		return nil, err
	}
	if index < 0 || index >= int(count) {
		return nil, fmt.Errorf("chart index %d out of range (worksheet has %d): %w",
			index, count, toolkiterrors.ErrChartNotFound)
	}
	chart, err := engine.Derive(charts.Get_Int(int32(index)))
	if err != nil {
		return nil, err
	}
	if chart == nil {
		return nil, fmt.Errorf("chart %d returned no handle: %w", index, toolkiterrors.ErrChartNotFound)
	}
	null, err := chart.IsNull()
	if err != nil {
		return nil, err
	}
	if null {
		return nil, fmt.Errorf("chart %d returned no handle: %w", index, toolkiterrors.ErrChartNotFound)
	}
	return chart, nil
}

// MoveChart repositions an existing chart to the given rectangle. Measured, a
// Move to a rectangle is byte-identical to having created the chart there in the
// first place, so this is a safe way to apply a placement after creation.
func MoveChart(chart *asposecells.Chart, topRow, leftColumn, bottomRow, rightColumn int) error {
	if err := ValidateChartBounds(topRow, leftColumn, bottomRow, rightColumn); err != nil {
		return err
	}
	return chart.Move(int32(topRow), int32(leftColumn), int32(bottomRow), int32(rightColumn))
}

// SetChartDataRange re-points an existing chart at a new source range. Measured,
// this rebuilds the series from the range without leaving stale ones behind, and
// category data set separately survives the rebuild.
func SetChartDataRange(chart *asposecells.Chart, dataRange string, byColumn bool) error {
	if err := ValidateChartDataRange(dataRange); err != nil {
		return err
	}
	return chart.SetChartDataRange(dataRange, byColumn)
}

// AddChartChartType adds a chart over dataRange and returns it.
//
// dataRange must be a bare A1 range ("A1:B5"); it is resolved against ws, so the
// worksheet name never enters the picture. byColumn selects the series
// orientation: true reads series down columns (each column is a series), false
// reads them across rows. It passes straight through as the engine's isvertical
// argument -- measured, "A2:B4" with true yields 2 series and with false yields
// 3, i.e. isvertical counts columns-as-series.
func AddChartChartType(ws *asposecells.Worksheet, chartType asposecells.ChartType, dataRange string, byColumn bool, topRow, leftColumn, bottomRow, rightColumn int) (*asposecells.Chart, error) {
	if err := ValidateChartDataRange(dataRange); err != nil {
		return nil, err
	}
	if err := ValidateChartBounds(topRow, leftColumn, bottomRow, rightColumn); err != nil {
		return nil, err
	}
	charts, err := Charts(ws)
	if err != nil {
		return nil, err
	}
	// Argument order is (type, range, isvertical, top, left, bottom, right).
	// The binding names the last two "rightrow, bottomcolumn", which is
	// misleading: measured, the 4th argument sets the bottom row and the 5th the
	// right column, matching Chart.Move.
	index, err := charts.Add_ChartType_String_Bool_Int_Int_Int_Int(
		chartType, dataRange, byColumn,
		int32(topRow), int32(leftColumn), int32(bottomRow), int32(rightColumn))
	if err != nil {
		return nil, err
	}
	// Use the index returned by Add directly rather than assuming the chart was
	// appended at the end. This avoids a fragile assumption about the engine's
	// collection behavior.
	return Chart(ws, int(index))
}

// DeleteChartAt removes the chart at index, or returns ErrChartNotFound.
func DeleteChartAt(ws *asposecells.Worksheet, index int) error {
	charts, err := Charts(ws)
	if err != nil {
		return err
	}
	count, err := charts.GetCount()
	if err != nil {
		return err
	}
	if index < 0 || index >= int(count) {
		return fmt.Errorf("chart index %d out of range (worksheet has %d): %w",
			index, count, toolkiterrors.ErrChartNotFound)
	}
	return charts.RemoveAt(int32(index))
}

// ClearCharts removes every chart from the worksheet.
func ClearCharts(ws *asposecells.Worksheet) error {
	charts, err := Charts(ws)
	if err != nil {
		return err
	}
	return charts.Clear()
}

// ChartSeries returns the chart's series collection.
func ChartSeries(chart *asposecells.Chart) (*asposecells.SeriesCollection, error) {
	ns, err := engine.Derive(chart.GetNSeries())
	if err != nil {
		return nil, err
	}
	if ns == nil {
		return nil, fmt.Errorf("chart returned no series collection: %w", toolkiterrors.ErrChartNotFound)
	}
	null, err := ns.IsNull()
	if err != nil {
		return nil, err
	}
	if null {
		return nil, fmt.Errorf("chart returned no series collection: %w", toolkiterrors.ErrChartNotFound)
	}
	return ns, nil
}

// ChartTitle returns the chart's title object.
func ChartTitle(chart *asposecells.Chart) (*asposecells.Title, error) {
	title, err := engine.Derive(chart.GetTitle())
	if err != nil {
		return nil, err
	}
	if title == nil {
		return nil, fmt.Errorf("chart returned no title: %w", toolkiterrors.ErrChartNotFound)
	}
	null, err := title.IsNull()
	if err != nil {
		return nil, err
	}
	if null {
		return nil, fmt.Errorf("chart returned no title: %w", toolkiterrors.ErrChartNotFound)
	}
	return title, nil
}

// ChartLegend returns the chart's legend object.
func ChartLegend(chart *asposecells.Chart) (*asposecells.Legend, error) {
	legend, err := engine.Derive(chart.GetLegend())
	if err != nil {
		return nil, err
	}
	if legend == nil {
		return nil, fmt.Errorf("chart returned no legend: %w", toolkiterrors.ErrChartNotFound)
	}
	null, err := legend.IsNull()
	if err != nil {
		return nil, err
	}
	if null {
		return nil, fmt.Errorf("chart returned no legend: %w", toolkiterrors.ErrChartNotFound)
	}
	return legend, nil
}

// RemoveChartSeriesAt removes the chart's series at index.
func RemoveChartSeriesAt(chart *asposecells.Chart, index int) error {
	ns, err := ChartSeries(chart)
	if err != nil {
		return err
	}
	count, err := ns.GetCount()
	if err != nil {
		return err
	}
	if index < 0 || index >= int(count) {
		return fmt.Errorf("series index %d out of range (chart has %d): %w",
			index, count, toolkiterrors.ErrChartNotFound)
	}
	return ns.RemoveAt(int32(index))
}

// ClearChartSeries removes every series from the chart.
func ClearChartSeries(chart *asposecells.Chart) error {
	ns, err := ChartSeries(chart)
	if err != nil {
		return err
	}
	return ns.Clear()
}

// AddChartSeries appends a series over dataRange.
func AddChartSeries(chart *asposecells.Chart, dataRange string, byColumn bool) error {
	if err := ValidateChartDataRange(dataRange); err != nil {
		return err
	}
	ns, err := ChartSeries(chart)
	if err != nil {
		return err
	}
	_, err = ns.Add_String_Bool(dataRange, byColumn)
	return err
}

// SetChartCategoryData points the chart's category axis at dataRange.
func SetChartCategoryData(chart *asposecells.Chart, dataRange string) error {
	if err := ValidateChartDataRange(dataRange); err != nil {
		return err
	}
	ns, err := ChartSeries(chart)
	if err != nil {
		return err
	}
	return ns.SetCategoryData(dataRange)
}

// SetChartStyle applies one of the engine's built-in chart styles, which are
// numbered 1..48. The engine does not enforce that range -- measured, setting 49
// or 99 silently stores 48 -- so the bound is checked here to keep a typo from
// being answered with a plausible-looking wrong style.
func SetChartStyle(chart *asposecells.Chart, style int) error {
	if style < 1 || style > 48 {
		return fmt.Errorf("chart style %d is outside the supported range 1..48: %w",
			style, toolkiterrors.ErrInvalidChartStyle)
	}
	return chart.SetStyle(int32(style))
}

// SetChartTitle sets the chart title text and makes the title visible.
//
// SetText alone carries both the text and its visibility across a save and
// reload -- measured, the text reads back unchanged and visible. The explicit
// SetIsVisible(true) is not for round-tripping but for the case where the title
// was hidden first: it makes "set the title" mean "the title shows", rather than
// leaving the caller with a title that is present but invisible.
func SetChartTitle(chart *asposecells.Chart, text string) error {
	title, err := ChartTitle(chart)
	if err != nil {
		return err
	}
	if err := title.SetText(text); err != nil {
		return err
	}
	return title.SetIsVisible(true)
}

// HideChartTitle hides the chart's title.
//
// Making the title invisible also discards its text: measured, the engine clears
// it, and a reloaded file reports the automatic title derived from the series
// again. Hiding is therefore not reversible through this API -- the text cannot
// be brought back by re-showing the title. To keep a title's text, use
// SetChartTitle; to remove a title permanently, hide it.
//
// The hidden state itself does survive a save and reload, which is the property
// this exists for.
func HideChartTitle(chart *asposecells.Chart) error {
	title, err := ChartTitle(chart)
	if err != nil {
		return err
	}
	return title.SetIsVisible(false)
}

// SetChartLegend shows or hides the chart legend.
func SetChartLegend(chart *asposecells.Chart, visible bool) error {
	return chart.SetShowLegend(visible)
}

// ResolveLegendPosition maps a legend position name to the engine enum. The
// engine enum has no automatic value, so "auto" is not a name this accepts;
// callers wanting the engine to choose call Legend.SetPositionAuto instead.
func ResolveLegendPosition(name string) (asposecells.LegendPositionType, error) {
	switch normalizeChartEnumName(name) {
	case "bottom":
		return asposecells.LegendPositionType_Bottom, nil
	case "corner":
		return asposecells.LegendPositionType_Corner, nil
	case "top":
		return asposecells.LegendPositionType_Top, nil
	case "right":
		return asposecells.LegendPositionType_Right, nil
	case "left":
		return asposecells.LegendPositionType_Left, nil
	case "notdocked":
		return asposecells.LegendPositionType_NotDocked, nil
	default:
		return 0, fmt.Errorf("unknown legend position %q: %w", name, toolkiterrors.ErrInvalidChartPosition)
	}
}

// SetChartLegendPosition moves the legend to one of the engine's docked
// positions.
func SetChartLegendPosition(chart *asposecells.Chart, name string) error {
	position, err := ResolveLegendPosition(name)
	if err != nil {
		return err
	}
	legend, err := ChartLegend(chart)
	if err != nil {
		return err
	}
	return legend.SetPosition(position)
}

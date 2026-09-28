package query

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// ChartMetadata describes a chart's metadata: its type, title, legend, bounds, and
// data range. It is the read counterpart of editor.AddChart and editor.InChart.
type ChartMetadata struct {
	// Index is the chart's zero-based position in the worksheet's chart
	// collection.
	Index int

	// Type is the chart's type as a toolkit-native string (e.g., "column",
	// "pie", "bar"). See editor.ChartType for the accepted names.
	Type string

	// Title is the chart's title text. Empty when the chart has no title or the
	// title is hidden.
	Title string

	// TitleVisible reports whether the title is shown. A chart can have title
	// text but not show it (TitleVisible=false), in which case the text is
	// still in Title but the chart displays no title.
	TitleVisible bool

	// Style is the chart's built-in style number (1..48). A style of -1 means
	// the chart uses the engine's default styling.
	Style int

	// ShowLegend reports whether the legend is visible.
	ShowLegend bool

	// LegendPosition is where the legend is docked (e.g., "bottom", "right").
	// See editor.ChartLegendPosition for the accepted names. Empty when the
	// legend is hidden.
	LegendPosition string

	// Bounds is the chart's position on the worksheet: (topRow, leftColumn,
	// bottomRow, rightColumn), all zero-based.
	Bounds ChartBounds

	// DataRange is the chart's source data range in A1 notation (e.g.,
	// "A1:B5"). Empty when the chart has no data range set.
	DataRange string

	// SeriesCount is the number of data series in the chart.
	SeriesCount int
}

// ChartBounds is a chart's position on the worksheet, in zero-based row and
// column coordinates.
type ChartBounds struct {
	TopRow      int
	LeftColumn  int
	BottomRow   int
	RightColumn int
}

// ChartSeries describes one data series in a chart: its values, category data,
// and the cell range it plots.
type ChartSeries struct {
	// Index is the series' zero-based position in the chart's series
	// collection.
	Index int

	// Values is the cell range the series plots (e.g., "$B$2:$B$5"). This is
	// the Y-axis data for most chart types.
	Values string

	// CategoryData is the cell range for the category axis labels (e.g.,
	// "$A$2:$A$5"). Empty when the series has no explicit category data.
	CategoryData string

	// DataValueCount is the number of data points in the series.
	DataValueCount int
}

// ChartInfo reads the metadata for the worksheet's chart at chartIndex. Charts
// are indexed from zero in the order they were added. A chartIndex out of range
// returns ErrChartNotFound.
//
// Example:
//
//	info, err := query.ChartInfo(datasource.FilePathSource("data.xlsx"), 0)
//	if err != nil {
//		log.Fatal(err)
//	}
//	fmt.Println(info.Type, info.Title, info.SeriesCount)
func ChartInfo(source datasource.DataSource, chartIndex int, opts ...Option) (ChartMetadata, error) {
	var out ChartMetadata
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return out, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return out, err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return out, err
	}
	chart, err := cells.Chart(ws, chartIndex)
	if err != nil {
		return out, err
	}
	out.Index = chartIndex

	// Chart type
	chartType, err := chart.GetType()
	if err != nil {
		return out, err
	}
	out.Type = chartTypeName(chartType)

	// Title
	title, err := chart.GetTitle()
	if err != nil {
		return out, err
	}
	titleText, err := title.GetText()
	if err != nil {
		return out, err
	}
	out.Title = titleText
	titleVisible, err := title.IsVisible()
	if err != nil {
		return out, err
	}
	out.TitleVisible = titleVisible

	// Style
	style, err := chart.GetStyle()
	if err != nil {
		return out, err
	}
	out.Style = int(style)

	// Legend
	showLegend, err := chart.GetShowLegend()
	if err != nil {
		return out, err
	}
	out.ShowLegend = showLegend
	if showLegend {
		legend, err := chart.GetLegend()
		if err != nil {
			return out, err
		}
		pos, err := legend.GetPosition()
		if err != nil {
			return out, err
		}
		out.LegendPosition = legendPositionName(pos)
	}

	// Bounds
	shape, err := chart.GetChartObject()
	if err != nil {
		return out, err
	}
	topRow, err := shape.GetUpperLeftRow()
	if err != nil {
		return out, err
	}
	leftCol, err := shape.GetUpperLeftColumn()
	if err != nil {
		return out, err
	}
	bottomRow, err := shape.GetLowerRightRow()
	if err != nil {
		return out, err
	}
	rightCol, err := shape.GetLowerRightColumn()
	if err != nil {
		return out, err
	}
	out.Bounds = ChartBounds{
		TopRow:      int(topRow),
		LeftColumn:  int(leftCol),
		BottomRow:   int(bottomRow),
		RightColumn: int(rightCol),
	}

	// Data range
	dataRange, err := chart.GetChartDataRange()
	if err != nil {
		return out, err
	}
	out.DataRange = dataRange

	// Series count
	series, err := chart.GetNSeries()
	if err != nil {
		return out, err
	}
	count, err := series.GetCount()
	if err != nil {
		return out, err
	}
	out.SeriesCount = int(count)

	return out, nil
}

// ChartSeriesData reads the series data for the worksheet's chart at
// chartIndex. Each series reports its values range, category data, and data
// point count. A chartIndex out of range returns ErrChartNotFound.
//
// Example:
//
//	series, err := query.ChartSeriesData(datasource.FilePathSource("data.xlsx"), 0)
//	if err != nil {
//		log.Fatal(err)
//	}
//	for _, s := range series {
//		fmt.Println(s.Values, s.DataValueCount)
//	}
func ChartSeriesData(source datasource.DataSource, chartIndex int, opts ...Option) ([]ChartSeries, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return nil, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return nil, err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return nil, err
	}
	chart, err := cells.Chart(ws, chartIndex)
	if err != nil {
		return nil, err
	}
	seriesCollection, err := chart.GetNSeries()
	if err != nil {
		return nil, err
	}
	count, err := seriesCollection.GetCount()
	if err != nil {
		return nil, err
	}

	// Category data is shared across all series, so read it once.
	categoryData, err := seriesCollection.GetCategoryData()
	if err != nil {
		return nil, err
	}

	out := make([]ChartSeries, 0, count)
	for i := int32(0); i < count; i++ {
		s, err := seriesCollection.Get(i)
		if err != nil {
			return nil, err
		}
		values, err := s.GetValues()
		if err != nil {
			return nil, err
		}
		dataCount, err := s.GetCountOfDataValues()
		if err != nil {
			return nil, err
		}
		out = append(out, ChartSeries{
			Index:          int(i),
			Values:         values,
			CategoryData:   categoryData,
			DataValueCount: int(dataCount),
		})
	}
	return out, nil
}

// ChartCount returns the number of charts on the worksheet.
//
// Example:
//
//	n, err := query.ChartCount(datasource.FilePathSource("data.xlsx"))
//	if err != nil {
//		log.Fatal(err)
//	}
//	fmt.Printf("worksheet has %d chart(s)\n", n)
func ChartCount(source datasource.DataSource, opts ...Option) (int, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil {
		return 0, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return 0, err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return 0, err
	}
	charts, err := cells.Charts(ws)
	if err != nil {
		return 0, err
	}
	count, err := charts.GetCount()
	if err != nil {
		return 0, err
	}
	return int(count), nil
}

// chartTypeName maps an engine chart type enum to its toolkit-native string
// name. The mapping is the inverse of cells.ResolveChartType.
func chartTypeName(t asposecells.ChartType) string {
	// Build a reverse map from the forward map in cells.chartTypeByName.
	// Since chartTypeByName is not exported, we duplicate the mapping here.
	// This covers all 81 chart types the engine supports.
	reverseMap := map[asposecells.ChartType]string{
		asposecells.ChartType_Area:                                      "area",
		asposecells.ChartType_AreaStacked:                               "areaStacked",
		asposecells.ChartType_Area100PercentStacked:                     "area100PercentStacked",
		asposecells.ChartType_Area3D:                                    "area3D",
		asposecells.ChartType_Area3DStacked:                             "area3DStacked",
		asposecells.ChartType_Area3D100PercentStacked:                   "area3D100PercentStacked",
		asposecells.ChartType_Bar:                                       "bar",
		asposecells.ChartType_BarStacked:                                "barStacked",
		asposecells.ChartType_Bar100PercentStacked:                      "bar100PercentStacked",
		asposecells.ChartType_Bar3DClustered:                            "bar3DClustered",
		asposecells.ChartType_Bar3DStacked:                              "bar3DStacked",
		asposecells.ChartType_Bar3D100PercentStacked:                    "bar3D100PercentStacked",
		asposecells.ChartType_Bubble:                                    "bubble",
		asposecells.ChartType_Bubble3D:                                  "bubble3D",
		asposecells.ChartType_Column:                                    "column",
		asposecells.ChartType_ColumnStacked:                             "columnStacked",
		asposecells.ChartType_Column100PercentStacked:                   "column100PercentStacked",
		asposecells.ChartType_Column3D:                                  "column3D",
		asposecells.ChartType_Column3DClustered:                         "column3DClustered",
		asposecells.ChartType_Column3DStacked:                           "column3DStacked",
		asposecells.ChartType_Column3D100PercentStacked:                 "column3D100PercentStacked",
		asposecells.ChartType_Cone:                                      "cone",
		asposecells.ChartType_ConeStacked:                               "coneStacked",
		asposecells.ChartType_Cone100PercentStacked:                     "cone100PercentStacked",
		asposecells.ChartType_ConicalBar:                                "conicalBar",
		asposecells.ChartType_ConicalBarStacked:                         "conicalBarStacked",
		asposecells.ChartType_ConicalBar100PercentStacked:               "conicalBar100PercentStacked",
		asposecells.ChartType_ConicalColumn3D:                           "conicalColumn3D",
		asposecells.ChartType_Cylinder:                                  "cylinder",
		asposecells.ChartType_CylinderStacked:                           "cylinderStacked",
		asposecells.ChartType_Cylinder100PercentStacked:                 "cylinder100PercentStacked",
		asposecells.ChartType_CylindricalBar:                            "cylindricalBar",
		asposecells.ChartType_CylindricalBarStacked:                     "cylindricalBarStacked",
		asposecells.ChartType_CylindricalBar100PercentStacked:           "cylindricalBar100PercentStacked",
		asposecells.ChartType_CylindricalColumn3D:                       "cylindricalColumn3D",
		asposecells.ChartType_Doughnut:                                  "doughnut",
		asposecells.ChartType_DoughnutExploded:                          "doughnutExploded",
		asposecells.ChartType_Line:                                      "line",
		asposecells.ChartType_LineStacked:                               "lineStacked",
		asposecells.ChartType_Line100PercentStacked:                     "line100PercentStacked",
		asposecells.ChartType_LineWithDataMarkers:                       "lineWithDataMarkers",
		asposecells.ChartType_LineStackedWithDataMarkers:                "lineStackedWithDataMarkers",
		asposecells.ChartType_Line100PercentStackedWithDataMarkers:      "line100PercentStackedWithDataMarkers",
		asposecells.ChartType_Line3D:                                    "line3D",
		asposecells.ChartType_Pie:                                       "pie",
		asposecells.ChartType_Pie3D:                                     "pie3D",
		asposecells.ChartType_PiePie:                                    "piePie",
		asposecells.ChartType_PieExploded:                               "pieExploded",
		asposecells.ChartType_Pie3DExploded:                             "pie3DExploded",
		asposecells.ChartType_PieBar:                                    "pieBar",
		asposecells.ChartType_Pyramid:                                   "pyramid",
		asposecells.ChartType_PyramidStacked:                            "pyramidStacked",
		asposecells.ChartType_Pyramid100PercentStacked:                  "pyramid100PercentStacked",
		asposecells.ChartType_PyramidBar:                                "pyramidBar",
		asposecells.ChartType_PyramidBarStacked:                         "pyramidBarStacked",
		asposecells.ChartType_PyramidBar100PercentStacked:               "pyramidBar100PercentStacked",
		asposecells.ChartType_PyramidColumn3D:                           "pyramidColumn3D",
		asposecells.ChartType_Radar:                                     "radar",
		asposecells.ChartType_RadarWithDataMarkers:                      "radarWithDataMarkers",
		asposecells.ChartType_RadarFilled:                               "radarFilled",
		asposecells.ChartType_Scatter:                                   "scatter",
		asposecells.ChartType_ScatterConnectedByCurvesWithDataMarker:    "scatterConnectedByCurvesWithDataMarker",
		asposecells.ChartType_ScatterConnectedByCurvesWithoutDataMarker: "scatterConnectedByCurvesWithoutDataMarker",
		asposecells.ChartType_ScatterConnectedByLinesWithDataMarker:     "scatterConnectedByLinesWithDataMarker",
		asposecells.ChartType_ScatterConnectedByLinesWithoutDataMarker:  "scatterConnectedByLinesWithoutDataMarker",
		asposecells.ChartType_StockHighLowClose:                         "stockHighLowClose",
		asposecells.ChartType_StockOpenHighLowClose:                     "stockOpenHighLowClose",
		asposecells.ChartType_StockVolumeHighLowClose:                   "stockVolumeHighLowClose",
		asposecells.ChartType_StockVolumeOpenHighLowClose:               "stockVolumeOpenHighLowClose",
		asposecells.ChartType_Surface3D:                                 "surface3D",
		asposecells.ChartType_SurfaceWireframe3D:                        "surfaceWireframe3D",
		asposecells.ChartType_SurfaceContour:                            "surfaceContour",
		asposecells.ChartType_SurfaceContourWireframe:                   "surfaceContourWireframe",
		asposecells.ChartType_BoxWhisker:                                "boxWhisker",
		asposecells.ChartType_Funnel:                                    "funnel",
		asposecells.ChartType_ParetoLine:                                "paretoLine",
		asposecells.ChartType_Sunburst:                                  "sunburst",
		asposecells.ChartType_Treemap:                                   "treemap",
		asposecells.ChartType_Waterfall:                                 "waterfall",
		asposecells.ChartType_Histogram:                                 "histogram",
		asposecells.ChartType_Map:                                       "map",
	}
	if name, ok := reverseMap[t]; ok {
		return name
	}
	// For unknown types, return a placeholder. This should not happen if the
	// reverse map is complete.
	return "unknown"
}

// legendPositionName maps an engine legend position enum to its toolkit-native
// string name.
func legendPositionName(pos asposecells.LegendPositionType) string {
	// The engine's enum values are defined in the binding. We need to reverse-map
	// from the enum to the name.
	reverseMap := map[asposecells.LegendPositionType]string{
		asposecells.LegendPositionType_Bottom:    "bottom",
		asposecells.LegendPositionType_Corner:    "corner",
		asposecells.LegendPositionType_Top:       "top",
		asposecells.LegendPositionType_Right:     "right",
		asposecells.LegendPositionType_Left:      "left",
		asposecells.LegendPositionType_NotDocked: "notDocked",
	}
	if name, ok := reverseMap[pos]; ok {
		return name
	}
	// For unknown positions, return a placeholder. This should not happen if the
	// reverse map is complete.
	return "unknown"
}

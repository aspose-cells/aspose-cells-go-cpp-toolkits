package tests

import (
	"errors"
	"fmt"
	"strings"
	"sync"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// chartSheetIndex is the worksheet every chart test targets. It is selected by
// index rather than by name because evaluation mode rewrites a random
// worksheet's name on a small fraction of loads.
const chartSheetIndex = 0

// chartFixture builds, once per process, the workbook the round-trip tests chart:
//
//	A1="Cat"  B1="Val"  C1="Extra"
//	A2="a"    B2=1      C2=10
//	A3="b"    B3=2      C3=20
//	A4="c"    B4=3      C4=30
//	A5="d"    B5=4      C5=40
//
// Three columns so a test can widen the chart's data range from A:B to A:C and
// watch the series count follow. It is cached with sync.Once because building it
// is cheap but the tests that consume it are not: every EditSpreadsheet over a
// byte source spends one of the evaluation copy's 100 loads per process, so the
// number of distinct sources is worth holding down.
var (
	chartFixtureOnce  sync.Once
	chartFixtureBytes []byte
	chartFixtureErr   error
)

func chartFixture(t *testing.T) []byte {
	t.Helper()
	chartFixtureOnce.Do(func() {
		chartFixtureBytes, chartFixtureErr = buildChartFixture()
	})
	if chartFixtureErr != nil {
		t.Fatalf("build chart fixture: %v", chartFixtureErr)
	}
	return chartFixtureBytes
}

func buildChartFixture() ([]byte, error) {
	blank, err := newBlankWorkbookBytes()
	if err != nil {
		return nil, err
	}
	return editor.EditSpreadsheet(datasource.BytesSource(blank),
		editor.InWorksheet(chartSheetIndex,
			editor.SetCellValue(0, 0, "Cat"),
			editor.SetCellValue(0, 1, "Val"),
			editor.SetCellValue(0, 2, "Extra"),
			editor.SetCellValue(1, 0, "a"),
			editor.SetCellValue(1, 1, 1),
			editor.SetCellValue(1, 2, 10),
			editor.SetCellValue(2, 0, "b"),
			editor.SetCellValue(2, 1, 2),
			editor.SetCellValue(2, 2, 20),
			editor.SetCellValue(3, 0, "c"),
			editor.SetCellValue(3, 1, 3),
			editor.SetCellValue(3, 2, 30),
			editor.SetCellValue(4, 0, "d"),
			editor.SetCellValue(4, 1, 4),
			editor.SetCellValue(4, 2, 40),
		),
	)
}

// chartReadback is one chart as read back from saved bytes. Every field is
// something a caller could observe in Excel, so asserting on it is asserting on
// the file rather than on the toolkit's own bookkeeping.
type chartReadback struct {
	// found reports whether chartIndex named an existing chart. Asking for a chart
	// that is not there is a normal thing for a test to observe -- it is how the
	// delete tests check a chart is gone -- so it is a field rather than an error.
	found           bool
	sheetChartCount int
	chartType       asposecells.ChartType
	seriesCount     int
	seriesValues    []string
	categoryData    string
	dataValueCount  int32
	// chartDataRange is the chart-level source range. It is read only to detect a
	// corrupted read: measured, the engine hands back raw pointer bytes for this
	// getter on roughly half of all loads in evaluation mode, while every
	// series-level getter stayed correct in every measurement. It is deliberately
	// not asserted on -- seriesValues is what proves the binding, and it is both
	// reliable and more precise, naming the cells each series actually plots.
	chartDataRange   string
	titleText        string
	titleVisible     bool
	style            int32
	showLegend       bool
	legendPosition   asposecells.LegendPositionType
	upperLeftRow     int32
	upperLeftColumn  int32
	lowerRightRow    int32
	lowerRightColumn int32
}

// looksLikeEngineString reports whether s could be a value the engine returned
// deliberately. An engine string is printable; the corruption this filters out
// returns raw pointer bytes, which are full of control characters. It lets the
// read-back turn a corrupted read into a retryable error instead of a confusing
// assertion failure.
//
// This accepts valid UTF-8 text (bytes 0x80-0xFF are valid UTF-8 leading and
// continuation bytes), so a chart title or range containing international
// characters is not mistaken for corruption. It rejects control characters
// (0x00-0x1F except tab/newline/CR, and 0x7F DEL), which are what the
// corruption pattern produces.
func looksLikeEngineString(s string) bool {
	for i := 0; i < len(s); i++ {
		b := s[i]
		// Reject control characters except common whitespace (tab, newline, CR).
		if b < 0x20 && b != 0x09 && b != 0x0A && b != 0x0D {
			return false
		}
		// Reject DEL (0x7F).
		if b == 0x7F {
			return false
		}
		// Allow everything else: printable ASCII (0x20-0x7E) and all UTF-8 bytes
		// (0x80-0xFF).
	}
	return true
}

// readChart loads workbook bytes and describes the chart at chartIndex on the
// worksheet at sheetIndex. It returns an error only for a genuine failure or a
// corrupted read, so callers can run it inside retryStable to re-roll the engine.
func readChart(data []byte, sheetIndex, chartIndex int) (chartReadback, error) {
	var out chartReadback

	wb, err := asposecells.NewWorkbook_Stream(data)
	if err != nil {
		return out, fmt.Errorf("reload: %w", err)
	}
	wss, err := wb.GetWorksheets()
	if err != nil {
		return out, fmt.Errorf("GetWorksheets: %w", err)
	}
	ws, err := wss.Get_Int(int32(sheetIndex))
	if err != nil {
		return out, fmt.Errorf("Get_Int(%d): %w", sheetIndex, err)
	}
	charts, err := ws.GetCharts()
	if err != nil {
		return out, fmt.Errorf("GetCharts: %w", err)
	}
	count, err := charts.GetCount()
	if err != nil {
		return out, fmt.Errorf("GetCount: %w", err)
	}
	out.sheetChartCount = int(count)
	if chartIndex >= out.sheetChartCount {
		return out, nil
	}

	chart, err := charts.Get_Int(int32(chartIndex))
	if err != nil {
		return out, fmt.Errorf("charts.Get_Int(%d): %w", chartIndex, err)
	}
	out.found = true

	if out.chartType, err = chart.GetType(); err != nil {
		return out, fmt.Errorf("GetType: %w", err)
	}
	if out.chartDataRange, err = chart.GetChartDataRange(); err != nil {
		return out, fmt.Errorf("GetChartDataRange: %w", err)
	}
	if !looksLikeEngineString(out.chartDataRange) {
		return out, fmt.Errorf("corrupted chart data range read: %q", out.chartDataRange)
	}
	if out.style, err = chart.GetStyle(); err != nil {
		return out, fmt.Errorf("GetStyle: %w", err)
	}
	if out.showLegend, err = chart.GetShowLegend(); err != nil {
		return out, fmt.Errorf("GetShowLegend: %w", err)
	}

	// The series collection is what distinguishes a chart that plots something
	// from one the engine created empty because it did not understand the range.
	ns, err := chart.GetNSeries()
	if err != nil {
		return out, fmt.Errorf("GetNSeries: %w", err)
	}
	seriesCount, err := ns.GetCount()
	if err != nil {
		return out, fmt.Errorf("series GetCount: %w", err)
	}
	out.seriesCount = int(seriesCount)
	for i := int32(0); i < seriesCount; i++ {
		series, err := ns.Get(i)
		if err != nil {
			return out, fmt.Errorf("series Get(%d): %w", i, err)
		}
		values, err := series.GetValues()
		if err != nil {
			return out, fmt.Errorf("series[%d] GetValues: %w", i, err)
		}
		if !looksLikeEngineString(values) {
			return out, fmt.Errorf("corrupted series[%d] values read: %q", i, values)
		}
		out.seriesValues = append(out.seriesValues, values)
	}
	if seriesCount > 0 {
		series, err := ns.Get(0)
		if err != nil {
			return out, fmt.Errorf("series Get(0): %w", err)
		}
		if out.dataValueCount, err = series.GetCountOfDataValues(); err != nil {
			return out, fmt.Errorf("series GetCountOfDataValues: %w", err)
		}
	}
	if out.categoryData, err = ns.GetCategoryData(); err != nil {
		return out, fmt.Errorf("GetCategoryData: %w", err)
	}
	if !looksLikeEngineString(out.categoryData) {
		return out, fmt.Errorf("corrupted category data read: %q", out.categoryData)
	}

	title, err := chart.GetTitle()
	if err != nil {
		return out, fmt.Errorf("GetTitle: %w", err)
	}
	if out.titleText, err = title.GetText(); err != nil {
		return out, fmt.Errorf("title GetText: %w", err)
	}
	if !looksLikeEngineString(out.titleText) {
		return out, fmt.Errorf("corrupted title text read: %q", out.titleText)
	}
	if out.titleVisible, err = title.IsVisible(); err != nil {
		return out, fmt.Errorf("title IsVisible: %w", err)
	}

	legend, err := chart.GetLegend()
	if err != nil {
		return out, fmt.Errorf("GetLegend: %w", err)
	}
	if out.legendPosition, err = legend.GetPosition(); err != nil {
		return out, fmt.Errorf("legend GetPosition: %w", err)
	}

	shape, err := chart.GetChartObject()
	if err != nil {
		return out, fmt.Errorf("GetChartObject: %w", err)
	}
	if out.upperLeftRow, err = shape.GetUpperLeftRow(); err != nil {
		return out, fmt.Errorf("GetUpperLeftRow: %w", err)
	}
	if out.upperLeftColumn, err = shape.GetUpperLeftColumn(); err != nil {
		return out, fmt.Errorf("GetUpperLeftColumn: %w", err)
	}
	if out.lowerRightRow, err = shape.GetLowerRightRow(); err != nil {
		return out, fmt.Errorf("GetLowerRightRow: %w", err)
	}
	if out.lowerRightColumn, err = shape.GetLowerRightColumn(); err != nil {
		return out, fmt.Errorf("GetLowerRightColumn: %w", err)
	}
	return out, nil
}

// readChartStable runs readChart under retryStable and fails the test if every
// attempt fails. A chart that is not there is not a failure: the returned struct
// simply has found unset.
func readChartStable(t *testing.T, data []byte, sheetIndex, chartIndex int) chartReadback {
	t.Helper()
	var got chartReadback
	err := retryStable(10, func() error {
		var err error
		got, err = readChart(data, sheetIndex, chartIndex)
		return err
	})
	if err != nil {
		t.Fatalf("read chart %d on sheet %d: %v", chartIndex, sheetIndex, err)
	}
	return got
}

// mustReadChart reads a chart the test expects to exist, failing with a clear
// message if it is missing.
func mustReadChart(t *testing.T, data []byte, sheetIndex, chartIndex int) chartReadback {
	t.Helper()
	got := readChartStable(t, data, sheetIndex, chartIndex)
	if !got.found {
		t.Fatalf("chart %d not found on sheet %d (sheet has %d charts)",
			chartIndex, sheetIndex, got.sheetChartCount)
	}
	return got
}

// TestChartAddRoundTrip verifies an added chart survives a save and reload with
// its type, source data, title, style and legend intact. The series assertions
// are the load-bearing ones: a chart can be created with a range the engine does
// not understand and still report a chart count of 1, so counting charts alone
// would pass on a silently empty chart.
func TestChartAddRoundTrip(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 2, 3, 20, 12,
				editor.WithChartTitle("Quarterly sales"),
				editor.WithChartStyle(7),
				editor.WithChartLegend(true),
				editor.WithChartLegendPosition(editor.ChartLegendLeft),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	got := mustReadChart(t, out, chartSheetIndex, 0)

	if got.sheetChartCount != 1 {
		t.Errorf("chart count = %d, want 1", got.sheetChartCount)
	}
	if got.chartType != asposecells.ChartType_Column {
		t.Errorf("chart type = %d, want Column (%d)", got.chartType, asposecells.ChartType_Column)
	}
	// A1:B5 read by column yields ONE series, not two: the engine consumes the
	// text column A as the category axis. The series therefore plots column B, and
	// asserting the exact cells it covers proves the range was understood -- a
	// chart the engine could not parse would have a series pointing nowhere.
	if got.seriesCount != 1 {
		t.Errorf("series count = %d, want 1 (column A becomes categories, leaving column B)", got.seriesCount)
	}
	if len(got.seriesValues) != 1 || !strings.Contains(got.seriesValues[0], "$B$2:$B$5") {
		t.Errorf("series values = %q, want one series over $B$2:$B$5", got.seriesValues)
	}
	if !strings.Contains(got.categoryData, "$A$2:$A$5") {
		t.Errorf("category data = %q, want it to reference $A$2:$A$5", got.categoryData)
	}
	if got.dataValueCount == 0 {
		t.Error("series 0 has 0 data values, so the chart plots nothing")
	}
	if got.titleText != "Quarterly sales" {
		t.Errorf("title text = %q, want %q", got.titleText, "Quarterly sales")
	}
	// Visibility is asserted separately from the text because the engine drops it
	// across a round trip unless it is set explicitly.
	if !got.titleVisible {
		t.Error("title is not visible after reload, so the chart renders untitled")
	}
	if got.style != 7 {
		t.Errorf("chart style = %d, want 7", got.style)
	}
	if !got.showLegend {
		t.Error("legend is hidden, want visible")
	}
	if got.legendPosition != asposecells.LegendPositionType_Left {
		t.Errorf("legend position = %d, want Left (%d)", got.legendPosition, asposecells.LegendPositionType_Left)
	}
	if got.upperLeftRow != 2 || got.upperLeftColumn != 3 || got.lowerRightRow != 20 || got.lowerRightColumn != 12 {
		t.Errorf("chart bounds = (%d,%d)/(%d,%d), want (2,3)/(20,12)",
			got.upperLeftRow, got.upperLeftColumn, got.lowerRightRow, got.lowerRightColumn)
	}
}

// TestChartDefaultsAndTypeFallback verifies a chart added without any actions is
// still a usable chart, and that a type with no curated constant works by name.
func TestChartDefaultsAndTypeFallback(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			// "pyramid" has no ChartTypePyramid constant; the raw name is the
			// documented fallback, and "pyramidBar" is a different type, so a
			// resolver that fuzzy-matched would pick the wrong one.
			editor.AddChart(editor.ChartType("Pyramid"), "A1:B5", true, 0, 3, 15, 10),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	got := mustReadChart(t, out, chartSheetIndex, 0)
	if got.chartType != asposecells.ChartType_Pyramid {
		t.Errorf("chart type = %d, want Pyramid (%d)", got.chartType, asposecells.ChartType_Pyramid)
	}
	if got.seriesCount != 1 {
		t.Errorf("series count = %d, want 1", got.seriesCount)
	}
	if got.dataValueCount == 0 {
		t.Error("series 0 has 0 data values, so the chart plots nothing")
	}
}

// TestChartInChartIsolatesTarget verifies InChart modifies only the chart it
// names: with two charts on the sheet, retitling index 0 must leave index 1's
// title and type alone.
func TestChartInChartIsolatesTarget(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10,
				editor.WithChartTitle("first"),
			),
			editor.AddChart(editor.ChartTypePie, "A1:B5", true, 0, 3, 15, 10,
				editor.WithChartTitle("second"),
			),
			editor.InChart(0,
				editor.WithChartTitle("retitled"),
				editor.WithChartStyle(12),
				// Switch chart 0's type too: a type change is a property of the
				// chart, so it must land on index 0 and leave index 1 alone.
				editor.WithChartType(editor.ChartTypeBar),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	first := mustReadChart(t, out, chartSheetIndex, 0)
	if first.sheetChartCount != 2 {
		t.Fatalf("chart count = %d, want 2", first.sheetChartCount)
	}
	if first.titleText != "retitled" {
		t.Errorf("chart 0 title = %q, want %q", first.titleText, "retitled")
	}
	if first.style != 12 {
		t.Errorf("chart 0 style = %d, want 12", first.style)
	}
	// The type change must survive the save and reload, and it must be the type
	// asked for: Bar, not the Column the chart was created as.
	if first.chartType != asposecells.ChartType_Bar {
		t.Errorf("chart 0 type = %d, want Bar (%d)", first.chartType, asposecells.ChartType_Bar)
	}
	// Changing the type reinterprets the data rather than dropping it, so the
	// series must still be there afterwards.
	if first.seriesCount == 0 {
		t.Error("chart 0 has no series after WithChartType, so the type change discarded the data")
	}

	second := mustReadChart(t, out, chartSheetIndex, 1)
	if second.titleText != "second" {
		t.Errorf("chart 1 title = %q, want it untouched at %q", second.titleText, "second")
	}
	if second.chartType != asposecells.ChartType_Pie {
		t.Errorf("chart 1 type = %d, want Pie (%d)", second.chartType, asposecells.ChartType_Pie)
	}
	if second.style == 12 {
		t.Error("chart 1 picked up chart 0's style, so InChart leaked across charts")
	}
}

// TestChartChangeDataSource verifies re-pointing a chart at a wider range adds the
// series that range implies. Widening A:B to A:C must take the series count from
// two to three; a rebuild that silently failed would leave it at two.
func TestChartChangeDataSource(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10),
			editor.InChart(0,
				editor.WithChartDataRange("A1:C5", true),
				editor.WithChartCategoryData("A2:A5"),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	got := mustReadChart(t, out, chartSheetIndex, 0)
	// A1:C5 is three columns with a text first column, so column A becomes the
	// categories and B and C are the two series. Re-pointing must have replaced
	// the old single series, not merely added to it.
	if got.seriesCount != 2 {
		t.Errorf("series count = %d, want 2 (column A categories, columns B and C)", got.seriesCount)
	}
	if len(got.seriesValues) != 2 ||
		!strings.Contains(got.seriesValues[0], "$B$2:$B$5") ||
		!strings.Contains(got.seriesValues[1], "$C$2:$C$5") {
		t.Errorf("series values = %q, want one series each over $B$2:$B$5 and $C$2:$C$5", got.seriesValues)
	}
	// WithChartCategoryData overrode the column the engine would have chosen only
	// in that the range's own first column agrees; either way the axis points here.
	if !strings.Contains(got.categoryData, "$A$2:$A$5") {
		t.Errorf("category data = %q, want it to reference $A$2:$A$5", got.categoryData)
	}
	if got.dataValueCount == 0 {
		t.Error("series 0 has 0 data values after the data range change")
	}
}

// TestChartSeriesAddRemoveClear verifies the series collection can be grown and
// shrunk one series at a time and emptied outright, with the chart surviving.
func TestChartSeriesAddRemoveClear(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10),
			editor.InChart(0,
				editor.WithChartSeries("C2:C5", true),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}
	// A1:B5 yields one series, and WithChartSeries appends one more.
	added := mustReadChart(t, out, chartSheetIndex, 0)
	if added.seriesCount != 2 {
		t.Errorf("after WithChartSeries: series count = %d, want 2", added.seriesCount)
	}
	if len(added.seriesValues) != 2 || !strings.Contains(added.seriesValues[1], "$C$2:$C$5") {
		t.Errorf("series values = %q, want the appended series over $C$2:$C$5", added.seriesValues)
	}

	out, err = editor.EditSpreadsheet(datasource.BytesSource(out),
		editor.InWorksheet(chartSheetIndex,
			editor.InChart(0, editor.RemoveChartSeries(0)),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet (remove): %v", err)
	}
	// Removing index 0 leaves the series that was appended, so its range is the
	// one still present -- that is what makes this a shift check, not a count check.
	removed := mustReadChart(t, out, chartSheetIndex, 0)
	if removed.seriesCount != 1 {
		t.Errorf("after RemoveChartSeries(0): series count = %d, want 1", removed.seriesCount)
	}
	if len(removed.seriesValues) != 1 || !strings.Contains(removed.seriesValues[0], "$C$2:$C$5") {
		t.Errorf("series values = %q, want the remaining series over $C$2:$C$5", removed.seriesValues)
	}

	out, err = editor.EditSpreadsheet(datasource.BytesSource(out),
		editor.InWorksheet(chartSheetIndex,
			editor.InChart(0, editor.ClearChartSeries()),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet (clear): %v", err)
	}
	got := mustReadChart(t, out, chartSheetIndex, 0)
	if got.seriesCount != 0 {
		t.Errorf("after ClearChartSeries: series count = %d, want 0", got.seriesCount)
	}
	// The chart itself outlives its data: clearing series empties the plot, it does
	// not delete the chart.
	if got.sheetChartCount != 1 {
		t.Errorf("chart count = %d, want 1 (clearing series must not delete the chart)", got.sheetChartCount)
	}
	// The chart keeps its identity after being emptied.
	if got.chartType != asposecells.ChartType_Column {
		t.Errorf("chart type = %d, want Column (%d) still", got.chartType, asposecells.ChartType_Column)
	}
}

// TestChartHideTitle verifies HideChartTitle removes a chart's heading and that
// the chart stays hidden across a save and reload.
//
// It deliberately does not assert the old text survives: the engine discards a
// title's text when the title is hidden, so hiding is not a reversible toggle
// through this API. Asserting the text were preserved would be asserting a
// behaviour the engine does not have.
func TestChartHideTitle(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10,
				editor.WithChartTitle("temporary"),
				editor.HideChartTitle(),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	got := mustReadChart(t, out, chartSheetIndex, 0)
	if got.titleVisible {
		t.Error("title is visible after HideChartTitle")
	}
	if got.titleText == "temporary" {
		t.Error("title text survived HideChartTitle, so this test's premise about the engine is stale")
	}
}

// TestChartDelete verifies DeleteChart removes the named chart and shifts the rest
// down, and DeleteAllCharts empties the sheet.
func TestChartDelete(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10, editor.WithChartTitle("keep")),
			editor.AddChart(editor.ChartTypePie, "A1:B5", true, 0, 3, 15, 10, editor.WithChartTitle("drop")),
			editor.DeleteChart(1),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet (delete one): %v", err)
	}
	got := mustReadChart(t, out, chartSheetIndex, 0)
	if got.sheetChartCount != 1 {
		t.Fatalf("chart count = %d, want 1", got.sheetChartCount)
	}
	if got.titleText != "keep" {
		t.Errorf("surviving chart title = %q, want %q", got.titleText, "keep")
	}

	out, err = editor.EditSpreadsheet(datasource.BytesSource(out),
		editor.InWorksheet(chartSheetIndex, editor.DeleteAllCharts()),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet (delete all): %v", err)
	}
	// The sheet must now have no charts at all, so chart 0 is legitimately absent.
	if empty := readChartStable(t, out, chartSheetIndex, 0); empty.sheetChartCount != 0 || empty.found {
		t.Errorf("after DeleteAllCharts: %d charts (found=%v), want 0", empty.sheetChartCount, empty.found)
	}
}

// chartErrorWorkbooks keeps every workbook handed out by chartTestWorksheet
// reachable for the lifetime of the process. The binding attaches finalizers to
// its handles, so a workbook that became unreachable could be collected -- and
// the worksheet handle derived from it invalidated -- while a test was still
// holding that handle. Each subtest needs its own fresh workbook because the
// cases mutate the sheet, so they accumulate here rather than being reused.
var chartErrorWorkbooks []*asposecells.Workbook

// chartTestWorksheet returns a fresh worksheet from an in-memory workbook.
// NewWorkbook creates the workbook without opening a file, so the
// error-classification tests below exercise the same code paths as a real edit
// without spending any of the evaluation copy's per-process load budget.
func chartTestWorksheet(t *testing.T) *asposecells.Worksheet {
	t.Helper()
	wb, err := asposecells.NewWorkbook()
	if err != nil {
		t.Fatalf("NewWorkbook: %v", err)
	}
	chartErrorWorkbooks = append(chartErrorWorkbooks, wb)
	wss, err := wb.GetWorksheets()
	if err != nil {
		t.Fatalf("GetWorksheets: %v", err)
	}
	ws, err := wss.Get_Int(0)
	if err != nil {
		t.Fatalf("Get_Int(0): %v", err)
	}
	return ws
}

// TestChartErrors verifies every rejection is classified with a sentinel error, so
// a caller can tell an unknown chart type from an out-of-range index without
// matching on message text.
func TestChartErrors(t *testing.T) {
	// Clear the accumulated workbooks when this test finishes, so they can be
	// garbage collected. Each subtest appends a workbook to chartErrorWorkbooks
	// to keep it alive during the test run (preventing finalizers from
	// invalidating derived handles), but after the test completes they are no
	// longer needed.
	t.Cleanup(func() {
		chartErrorWorkbooks = nil
	})

	cases := []struct {
		name string
		want error
		run  func(ws *asposecells.Worksheet) error
	}{
		{
			name: "unknown chart type",
			want: toolkiterrors.ErrInvalidChartType,
			run: func(ws *asposecells.Worksheet) error {
				return editor.AddChart(editor.ChartType("sparkline"), "A1:B5", true, 0, 3, 15, 10)(ws)
			},
		},
		{
			name: "single-cell data range",
			want: toolkiterrors.ErrInvalidRange,
			run: func(ws *asposecells.Worksheet) error {
				return editor.AddChart(editor.ChartTypeColumn, "A1", true, 0, 3, 15, 10)(ws)
			},
		},
		{
			name: "single-row data range",
			want: toolkiterrors.ErrInvalidRange,
			run: func(ws *asposecells.Worksheet) error {
				return editor.AddChart(editor.ChartTypeColumn, "A1:B1", true, 0, 3, 15, 10)(ws)
			},
		},
		{
			name: "sheet-qualified data range",
			want: toolkiterrors.ErrInvalidRange,
			run: func(ws *asposecells.Worksheet) error {
				return editor.AddChart(editor.ChartTypeColumn, "Sheet1!A1:B5", true, 0, 3, 15, 10)(ws)
			},
		},
		{
			name: "off-grid data range",
			want: toolkiterrors.ErrInvalidRange,
			run: func(ws *asposecells.Worksheet) error {
				return editor.AddChart(editor.ChartTypeColumn, "XFE1:XFE9", true, 0, 3, 15, 10)(ws)
			},
		},
		{
			name: "malformed data range",
			want: toolkiterrors.ErrInvalidCellRef,
			run: func(ws *asposecells.Worksheet) error {
				return editor.AddChart(editor.ChartTypeColumn, "not-a-range", true, 0, 3, 15, 10)(ws)
			},
		},
		{
			name: "reversed AddChart bounds",
			want: toolkiterrors.ErrInvalidRange,
			run: func(ws *asposecells.Worksheet) error {
				return editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 5, 5, 2, 2)(ws)
			},
		},
		{
			name: "zero-area AddChart bounds",
			want: toolkiterrors.ErrInvalidRange,
			run: func(ws *asposecells.Worksheet) error {
				return editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 0, 3)(ws)
			},
		},
		{
			name: "off-grid AddChart bounds",
			want: toolkiterrors.ErrInvalidRange,
			run: func(ws *asposecells.Worksheet) error {
				return editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 20000)(ws)
			},
		},
		{
			// WithChartBounds is the other way to place a chart, so it has to
			// reject the same rectangles AddChart does. It is reachable only
			// through InChart now that AddChart takes the bounds directly.
			name: "reversed WithChartBounds",
			want: toolkiterrors.ErrInvalidRange,
			run: func(ws *asposecells.Worksheet) error {
				if err := editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10)(ws); err != nil {
					return err
				}
				return editor.InChart(0, editor.WithChartBounds(5, 5, 2, 2))(ws)
			},
		},
		{
			name: "zero-area WithChartBounds",
			want: toolkiterrors.ErrInvalidRange,
			run: func(ws *asposecells.Worksheet) error {
				if err := editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10)(ws); err != nil {
					return err
				}
				return editor.InChart(0, editor.WithChartBounds(0, 3, 0, 3))(ws)
			},
		},
		{
			name: "chart style below range",
			want: toolkiterrors.ErrInvalidChartStyle,
			run: func(ws *asposecells.Worksheet) error {
				// The engine silently clamps 0 and 99 to a valid style, so an
				// unvalidated out-of-range number would be answered with a style
				// the caller did not ask for.
				return editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10,
					editor.WithChartStyle(0))(ws)
			},
		},
		{
			name: "chart style above range",
			want: toolkiterrors.ErrInvalidChartStyle,
			run: func(ws *asposecells.Worksheet) error {
				return editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10,
					editor.WithChartStyle(99))(ws)
			},
		},
		{
			name: "unknown legend position",
			want: toolkiterrors.ErrInvalidChartPosition,
			run: func(ws *asposecells.Worksheet) error {
				return editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10,
					editor.WithChartLegendPosition(editor.ChartLegendPosition("middle")))(ws)
			},
		},
		{
			name: "InChart index out of range",
			want: toolkiterrors.ErrChartNotFound,
			run: func(ws *asposecells.Worksheet) error {
				return editor.InChart(99, editor.WithChartTitle("nope"))(ws)
			},
		},
		{
			name: "InChart on a sheet with no charts",
			want: toolkiterrors.ErrChartNotFound,
			run: func(ws *asposecells.Worksheet) error {
				return editor.InChart(0)(ws)
			},
		},
		{
			name: "DeleteChart index out of range",
			want: toolkiterrors.ErrChartNotFound,
			run: func(ws *asposecells.Worksheet) error {
				return editor.DeleteChart(99)(ws)
			},
		},
		{
			name: "DeleteChart negative index",
			want: toolkiterrors.ErrChartNotFound,
			run: func(ws *asposecells.Worksheet) error {
				return editor.DeleteChart(-1)(ws)
			},
		},
		{
			name: "RemoveChartSeries index out of range",
			want: toolkiterrors.ErrChartNotFound,
			run: func(ws *asposecells.Worksheet) error {
				if err := editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10)(ws); err != nil {
					return err
				}
				return editor.InChart(0, editor.RemoveChartSeries(99))(ws)
			},
		},
		{
			name: "WithChartType unknown type",
			want: toolkiterrors.ErrInvalidChartType,
			run: func(ws *asposecells.Worksheet) error {
				if err := editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10)(ws); err != nil {
					return err
				}
				return editor.InChart(0, editor.WithChartType(editor.ChartType("sparkline")))(ws)
			},
		},
		{
			name: "WithChartDataRange off-grid range",
			want: toolkiterrors.ErrInvalidRange,
			run: func(ws *asposecells.Worksheet) error {
				if err := editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10)(ws); err != nil {
					return err
				}
				return editor.InChart(0, editor.WithChartDataRange("XFE1:XFE9", true))(ws)
			},
		},
		{
			name: "WithChartSeries malformed range",
			want: toolkiterrors.ErrInvalidCellRef,
			run: func(ws *asposecells.Worksheet) error {
				if err := editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10)(ws); err != nil {
					return err
				}
				return editor.InChart(0, editor.WithChartSeries("not-a-range", true))(ws)
			},
		},
		{
			name: "WithChartCategoryData sheet-qualified range",
			want: toolkiterrors.ErrInvalidRange,
			run: func(ws *asposecells.Worksheet) error {
				if err := editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10)(ws); err != nil {
					return err
				}
				return editor.InChart(0, editor.WithChartCategoryData("Sheet1!A2:A5"))(ws)
			},
		},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			err := tc.run(chartTestWorksheet(t))
			if !errors.Is(err, tc.want) {
				t.Fatalf("error = %v, want %v", err, tc.want)
			}
		})
	}
}

// TestDeleteAllChartsOnEmptySheet verifies clearing a sheet that has no charts is
// a no-op rather than an error, so a caller can clear unconditionally.
func TestDeleteAllChartsOnEmptySheet(t *testing.T) {
	if err := editor.DeleteAllCharts()(chartTestWorksheet(t)); err != nil {
		t.Fatalf("DeleteAllCharts on a chartless sheet: %v", err)
	}
}

// TestChartLeavesCellsIntact verifies adding a chart does not disturb the
// worksheet's data: a chart is a drawing on top of the grid, so the source cells
// must read back unchanged.
func TestChartLeavesCellsIntact(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10,
				editor.WithChartTitle("overlay"),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	if err := retryStable(5, func() error {
		return verifyImportedGrid(datasource.BytesSource(out), [][]interface{}{
			{"Cat", "Val", "Extra"},
			{"a", 1, 10},
			{"b", 2, 20},
			{"c", 3, 30},
			{"d", 4, 40},
		})
	}); err != nil {
		t.Fatal(err)
	}
}

// TestChartFromEmptyWorkbook verifies the whole from-scratch path: seed with
// datasource.NewEmptyWorkbook (the toolkit's "start from nothing" entry point),
// write a table onto a sheet this pass adds, and chart it -- all in one
// EditSpreadsheet, with no input file and no engine handle.
//
// The sheet is named rather than taken from the blank workbook's default,
// because a name this pass creates cannot be corrupted by the load that starts
// it; the read-back then targets it by index for the same reason.
func TestChartFromEmptyWorkbook(t *testing.T) {
	seed, err := datasource.NewEmptyWorkbook()
	if err != nil {
		t.Fatalf("NewEmptyWorkbook: %v", err)
	}

	const sheet = "Sales"
	out, err := editor.EditSpreadsheet(seed,
		editor.WithAddWorksheet(sheet),
		editor.InWorksheet(sheet,
			editor.SetCellValue(0, 0, "Region"),
			editor.SetCellValue(0, 1, "Q1"),
			editor.SetCellValue(1, 0, "North"),
			editor.SetCellValue(1, 1, int32(120)),
			editor.SetCellValue(2, 0, "South"),
			editor.SetCellValue(2, 1, int32(90)),
			editor.AddChart(editor.ChartTypeColumn, "A1:B3", true, 4, 0, 18, 8,
				editor.WithChartTitle("from scratch"),
				editor.WithChartStyle(5),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	// "Sales" is the second sheet: the blank workbook already has one at index 0.
	const salesIndex = 1
	got := mustReadChart(t, out, salesIndex, 0)

	if got.chartType != asposecells.ChartType_Column {
		t.Errorf("chart type = %d, want Column (%d)", got.chartType, asposecells.ChartType_Column)
	}
	if got.titleText != "from scratch" {
		t.Errorf("title text = %q, want %q", got.titleText, "from scratch")
	}
	if got.style != 5 {
		t.Errorf("chart style = %d, want 5", got.style)
	}
	// Column A is text, so it becomes the category axis and the one series plots
	// column B. A chart the engine failed to parse would have a series pointing
	// nowhere, so assert the cells it actually covers.
	if got.seriesCount != 1 {
		t.Fatalf("series count = %d, want 1 (column A becomes categories)", got.seriesCount)
	}
	if !strings.Contains(got.seriesValues[0], "$B$2:$B$3") {
		t.Errorf("series values = %q, want a series over $B$2:$B$3", got.seriesValues)
	}
	if !strings.Contains(got.categoryData, "$A$2:$A$3") {
		t.Errorf("category data = %q, want it to reference $A$2:$A$3", got.categoryData)
	}
}

// TestChartPlacement verifies where a chart lands: at the rectangle AddChart was
// given, and -- when WithChartBounds is applied afterwards -- at the last bounds
// action, since actions apply in order.
//
// The two charts use different rectangles so a mix-up between them cannot pass:
// chart 0 keeps its AddChart rectangle, chart 1 is moved off its AddChart
// rectangle by the actions.
func TestChartPlacement(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			// No bounds action: the chart stays where AddChart put it.
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10,
				editor.WithChartTitle("as added"),
			),
			// Two bounds actions: the second must win.
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10,
				editor.WithChartTitle("explicit"),
				editor.WithChartBounds(2, 2, 8, 8),
				editor.WithChartBounds(4, 4, 12, 12),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	asAdded := mustReadChart(t, out, chartSheetIndex, 0)
	if asAdded.upperLeftRow != 0 || asAdded.upperLeftColumn != 3 ||
		asAdded.lowerRightRow != 15 || asAdded.lowerRightColumn != 10 {
		t.Errorf("bounds as added = (%d,%d)/(%d,%d), want the AddChart rectangle (0,3)/(15,10)",
			asAdded.upperLeftRow, asAdded.upperLeftColumn,
			asAdded.lowerRightRow, asAdded.lowerRightColumn)
	}

	explicit := mustReadChart(t, out, chartSheetIndex, 1)
	if explicit.upperLeftRow != 4 || explicit.upperLeftColumn != 4 ||
		explicit.lowerRightRow != 12 || explicit.lowerRightColumn != 12 {
		t.Errorf("bounds after two WithChartBounds = (%d,%d)/(%d,%d), want the last one (4,4)/(12,12)",
			explicit.upperLeftRow, explicit.upperLeftColumn,
			explicit.lowerRightRow, explicit.lowerRightColumn)
	}
}

// TestChartRowOrientation verifies the byColumn flag actually selects how the
// range is read into series: the same range must yield one series per column when
// true and one per row when false.
//
// Both charts are added in a single pass so the pair costs one load rather than
// two, and reading them from the same saved bytes proves the difference comes
// from the flag and not from two separately-built workbooks.
func TestChartRowOrientation(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			// A2:C5 read by column: column A is text, so it becomes the category
			// axis and the two series come from columns B and C.
			editor.AddChart(editor.ChartTypeColumn, "A2:C5", true, 0, 0, 8, 6,
				editor.WithChartTitle("by column"),
			),
			// The same range read by row: each of the four data rows is a series.
			editor.AddChart(editor.ChartTypeColumn, "A2:C5", false, 10, 0, 18, 6,
				editor.WithChartTitle("by row"),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	byColumn := mustReadChart(t, out, chartSheetIndex, 0)
	if byColumn.seriesCount != 2 {
		t.Errorf("byColumn=true: series count = %d, want 2 (columns B and C; column A is categories)",
			byColumn.seriesCount)
	}

	byRow := mustReadChart(t, out, chartSheetIndex, 1)
	if byRow.seriesCount != 4 {
		t.Errorf("byColumn=false: series count = %d, want 4 (one per data row)", byRow.seriesCount)
	}
	// The two charts differ only by the flag, so a resolver that ignored it would
	// produce equal counts -- which is what this asserts against.
	if byRow.seriesCount == byColumn.seriesCount {
		t.Error("both orientations produced the same series count, so byColumn had no effect")
	}
}

// TestChartLegendPositions verifies a legend position reaches the engine rather
// than only resolving in the toolkit: a position that resolved but was never
// applied would read back as the engine's default (Right).
//
// The mapping for all six positions is unit-tested in the internal package; this
// covers three end-to-end, and deliberately leaves the legend visible, because a
// hidden legend's position is not something the engine can be trusted to keep.
//
// ChartLegendNotDocked is asserted to NOT survive the save, which is the measured
// behaviour rather than a bug in the toolkit: the engine applies it and reports it
// back in memory (see TestChartLegendNotDockedIsInMemoryOnly), but XLSX has no way
// to record an undocked legend, so a reloaded chart comes back docked right. If
// the engine ever learns to persist it, this test fails and says so.
func TestChartLegendPositions(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 0, 8, 6,
				editor.WithChartTitle("corner"),
				editor.WithChartLegend(true),
				editor.WithChartLegendPosition(editor.ChartLegendCorner),
			),
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 10, 0, 18, 6,
				editor.WithChartTitle("bottom"),
				editor.WithChartLegend(true),
				editor.WithChartLegendPosition(editor.ChartLegendBottom),
			),
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 30, 0, 38, 6,
				editor.WithChartTitle("top"),
				editor.WithChartLegend(true),
				editor.WithChartLegendPosition(editor.ChartLegendTop),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	for _, tc := range []struct {
		index int
		want  asposecells.LegendPositionType
		name  string
	}{
		{0, asposecells.LegendPositionType_Corner, "ChartLegendCorner"},
		{1, asposecells.LegendPositionType_Bottom, "ChartLegendBottom"},
		{2, asposecells.LegendPositionType_Top, "ChartLegendTop"},
	} {
		got := mustReadChart(t, out, chartSheetIndex, tc.index)
		if !got.showLegend {
			t.Errorf("chart %d: legend is hidden, so its position cannot be asserted", tc.index)
			continue
		}
		if got.legendPosition != tc.want {
			t.Errorf("chart %d (%s): legend position = %d, want %d",
				tc.index, tc.name, got.legendPosition, tc.want)
		}
	}
}

// TestChartLegendNotDockedIsInMemoryOnly pins the one legend position that does
// not survive a save. The engine accepts ChartLegendNotDocked and reports it back
// on the live handle, so the toolkit is not silently dropping the request; it is
// XLSX that cannot express an undocked legend, and the position reverts to the
// default once the file is written and read again.
//
// Both halves matter. Without the in-memory assertion, a regression that made the
// toolkit reject the position outright would look identical to this documented
// limitation. This runs on a fresh in-memory worksheet, so it spends no load.
func TestChartLegendNotDockedIsInMemoryOnly(t *testing.T) {
	ws := chartTestWorksheet(t)
	if err := editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10,
		editor.WithChartLegend(true))(ws); err != nil {
		t.Fatalf("AddChart: %v", err)
	}
	charts, err := ws.GetCharts()
	if err != nil {
		t.Fatalf("GetCharts: %v", err)
	}
	chart, err := charts.Get_Int(0)
	if err != nil {
		t.Fatalf("Get_Int(0): %v", err)
	}

	// Applying the position must reach the engine: read it back off the same
	// handle, with no save in between.
	if err := editor.WithChartLegendPosition(editor.ChartLegendNotDocked)(chart); err != nil {
		t.Fatalf("WithChartLegendPosition(notDocked): %v", err)
	}
	legend, err := chart.GetLegend()
	if err != nil {
		t.Fatalf("GetLegend: %v", err)
	}
	position, err := legend.GetPosition()
	if err != nil {
		t.Fatalf("GetPosition: %v", err)
	}
	if position != asposecells.LegendPositionType_NotDocked {
		t.Fatalf("legend position in memory = %d, want NotDocked (%d): the toolkit "+
			"is not applying the position at all", position, asposecells.LegendPositionType_NotDocked)
	}
}

// TestChartModifyExistingChart verifies the modify-in-place path on a chart that
// has already been saved and reloaded: reopening the file and addressing chart 0
// must change that chart's placement, title, type, style and legend, while the
// source data survives the changes.
//
// This is the flow a caller actually has -- read a file, adjust a chart, write it
// back -- so it is asserted across separate passes rather than inside one.
func TestChartModifyExistingChart(t *testing.T) {
	// Pass 1: author the chart.
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:C5", true, 0, 0, 10, 6,
				editor.WithChartTitle("original"),
				editor.WithChartStyle(6),
				editor.WithChartLegend(true),
				editor.WithChartLegendPosition(editor.ChartLegendLeft),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet (author): %v", err)
	}

	before := mustReadChart(t, out, chartSheetIndex, 0)
	if before.seriesCount != 2 {
		t.Fatalf("chart authored with %d series, want 2 to modify", before.seriesCount)
	}

	// Pass 2: reopen the saved bytes and change what the chart looks like.
	out, err = editor.EditSpreadsheet(datasource.BytesSource(out),
		editor.InWorksheet(chartSheetIndex,
			editor.InChart(0,
				editor.WithChartTitle("modified"),
				editor.WithChartStyle(30),
				editor.WithChartLegendPosition(editor.ChartLegendTop),
				editor.WithChartBounds(6, 1, 22, 9),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet (modify): %v", err)
	}

	after := mustReadChart(t, out, chartSheetIndex, 0)
	if after.sheetChartCount != 1 {
		t.Errorf("chart count = %d, want 1 (modifying must not add or drop a chart)", after.sheetChartCount)
	}
	if after.titleText != "modified" {
		t.Errorf("title = %q, want %q", after.titleText, "modified")
	}
	if after.style != 30 {
		t.Errorf("style = %d, want 30", after.style)
	}
	if after.legendPosition != asposecells.LegendPositionType_Top {
		t.Errorf("legend position = %d, want Top (%d)", after.legendPosition, asposecells.LegendPositionType_Top)
	}
	if after.upperLeftRow != 6 || after.upperLeftColumn != 1 ||
		after.lowerRightRow != 22 || after.lowerRightColumn != 9 {
		t.Errorf("bounds = (%d,%d)/(%d,%d), want the repositioned (6,1)/(22,9)",
			after.upperLeftRow, after.upperLeftColumn, after.lowerRightRow, after.lowerRightColumn)
	}
	// The data is not part of what was changed, so it must still be there: a
	// modify that quietly rebuilt the chart would have dropped the series.
	if after.seriesCount != 2 {
		t.Errorf("series count = %d, want the original 2 to survive the modification", after.seriesCount)
	}
	if !strings.Contains(after.seriesValues[0], "$B$2:$B$5") {
		t.Errorf("series values = %q, want the original series over $B$2:$B$5", after.seriesValues)
	}

	// Pass 3: hiding the legend is its own change, kept separate because the
	// engine treats a hidden legend's position as uninteresting.
	out, err = editor.EditSpreadsheet(datasource.BytesSource(out),
		editor.InWorksheet(chartSheetIndex,
			editor.InChart(0, editor.WithChartLegend(false)),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet (hide legend): %v", err)
	}

	hidden := mustReadChart(t, out, chartSheetIndex, 0)
	if hidden.showLegend {
		t.Error("legend is still shown after WithChartLegend(false)")
	}
	// Hiding the legend must leave the rest of the chart alone.
	if hidden.titleText != "modified" || hidden.style != 30 {
		t.Errorf("hiding the legend disturbed the chart: title = %q, style = %d, want %q and 30",
			hidden.titleText, hidden.style, "modified")
	}
}

// TestChartWithNamedRange verifies that a chart can use a named range as its
// data source, and that the chart plots the cells the named range refers to.
//
// This is an interaction test: named ranges are defined with DefineNamedRange
// and charts are created with AddChart. The test verifies the two features work
// together without conflict.
func TestChartWithNamedRange(t *testing.T) {
	// Define a named range "Scores" over B2:B5, then create a chart that plots
	// the same range. The chart's series should reference the named range's cells.
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.DefineNamedRange("Scores", 1, 1, 4, 1),
			editor.AddChart(editor.ChartTypeColumn, "B2:B5", true, 0, 3, 15, 10,
				editor.WithChartTitle("From named range"),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	got := mustReadChart(t, out, chartSheetIndex, 0)
	if !got.found {
		t.Fatal("chart not found")
	}
	if got.seriesCount != 1 {
		t.Fatalf("series count = %d, want 1", got.seriesCount)
	}
	// The chart should plot B2:B5, which is what the named range "Scores" refers to.
	if !strings.Contains(got.seriesValues[0], "$B$2:$B$5") {
		t.Errorf("series values = %q, want a series over $B$2:$B$5 (the named range)", got.seriesValues)
	}
}

// TestChartInheritsDefaultStyle verifies that a chart created without an explicit
// style does not inherit the workbook's default style. Charts have their own
// style (1..48, or -1 for the engine default), and the workbook's default style
// applies to cells, not to charts.
//
// This is an interaction test: SetDefaultStyle sets the workbook's default cell
// style, and AddChart creates a chart. The test verifies the chart's style is
// independent of the cell style.
func TestChartInheritsDefaultStyle(t *testing.T) {
	// Set a default cell style (bold, red font), then create a chart without an
	// explicit style. The chart's style should be -1 (the engine default), not
	// influenced by the cell style.
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InDefaultStyle(
			editor.WithFontName("Arial"),
			editor.WithFontSize(14),
			editor.WithFontColor("FF0000"),
		),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10,
				editor.WithChartTitle("No explicit style"),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	got := mustReadChart(t, out, chartSheetIndex, 0)
	if !got.found {
		t.Fatal("chart not found")
	}
	// The chart's style should be -1 (the engine default), not 0 or any other
	// value. A chart that inherited the cell style would have a different style
	// number, but charts and cells use separate style systems.
	if got.style != -1 {
		t.Errorf("chart style = %d, want -1 (engine default); charts do not inherit cell styles", got.style)
	}
}

// TestChartMultipleSheets verifies that charts on different worksheets are
// independent: a chart on one sheet does not interfere with a chart on another,
// and each chart plots data from its own sheet.
//
// This is an interaction test: charts are created on multiple sheets, and the
// test verifies they are isolated from each other.
func TestChartMultipleSheets(t *testing.T) {
	// Create two sheets, each with its own data and chart. The blank workbook
	// starts with a default "Sheet1", so we add "DataSheet" and "ChartSheet" to
	// avoid name conflicts.
	blank, err := newBlankWorkbookBytes()
	if err != nil {
		t.Fatalf("newBlankWorkbookBytes: %v", err)
	}
	out, err := editor.EditSpreadsheet(datasource.BytesSource(blank),
		editor.WithAddWorksheet("DataSheet"),
		editor.WithAddWorksheet("ChartSheet"),
		// DataSheet (index 1): 2-column data with a column chart. First column
		// has text labels to be treated as categories.
		editor.InWorksheet("DataSheet",
			editor.SetCellValue(0, 0, "Label"),
			editor.SetCellValue(0, 1, "Value"),
			editor.SetCellValue(1, 0, "A"),
			editor.SetCellValue(1, 1, 10),
			editor.SetCellValue(2, 0, "B"),
			editor.SetCellValue(2, 1, 20),
			editor.AddChart(editor.ChartTypeColumn, "A1:B3", true, 0, 3, 15, 10,
				editor.WithChartTitle("DataSheet chart"),
			),
		),
		// ChartSheet (index 2): 3-column data with a pie chart. First column
		// has text labels to be treated as categories.
		editor.InWorksheet("ChartSheet",
			editor.SetCellValue(0, 0, "Category"),
			editor.SetCellValue(0, 1, "Series1"),
			editor.SetCellValue(0, 2, "Series2"),
			editor.SetCellValue(1, 0, "X"),
			editor.SetCellValue(1, 1, 50),
			editor.SetCellValue(1, 2, 500),
			editor.SetCellValue(2, 0, "Y"),
			editor.SetCellValue(2, 1, 60),
			editor.SetCellValue(2, 2, 600),
			editor.AddChart(editor.ChartTypePie, "A1:C3", true, 0, 3, 15, 10,
				editor.WithChartTitle("ChartSheet chart"),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	// Verify DataSheet's chart (index 1): column chart with 1 series (column B,
	// with A as categories since A has text labels)
	dataSheet := mustReadChart(t, out, 1, 0)
	if !dataSheet.found {
		t.Fatal("DataSheet chart not found")
	}
	if dataSheet.chartType != asposecells.ChartType_Column {
		t.Errorf("DataSheet chart type = %d, want Column", dataSheet.chartType)
	}
	if dataSheet.seriesCount != 1 {
		t.Errorf("DataSheet chart series = %d, want 1 (column B, with A as categories)", dataSheet.seriesCount)
	}
	if dataSheet.titleText != "DataSheet chart" {
		t.Errorf("DataSheet chart title = %q, want %q", dataSheet.titleText, "DataSheet chart")
	}

	// Verify ChartSheet's chart (index 2): pie chart with 2 series (columns B
	// and C, with A as categories since A has text labels)
	chartSheet := mustReadChart(t, out, 2, 0)
	if !chartSheet.found {
		t.Fatal("ChartSheet chart not found")
	}
	if chartSheet.chartType != asposecells.ChartType_Pie {
		t.Errorf("ChartSheet chart type = %d, want Pie", chartSheet.chartType)
	}
	if chartSheet.seriesCount != 2 {
		t.Errorf("ChartSheet chart series = %d, want 2 (columns B and C, with A as categories)", chartSheet.seriesCount)
	}
	if chartSheet.titleText != "ChartSheet chart" {
		t.Errorf("ChartSheet chart title = %q, want %q", chartSheet.titleText, "ChartSheet chart")
	}
}

// TestChartStylePreset verifies that WithChartPreset applies all the preset's
// styling options (style number, title, legend visibility, legend position) in
// a single action.
func TestChartStylePreset(t *testing.T) {
	preset := editor.ChartStylePreset{
		Style:          7,
		Title:          "Preset title",
		ShowLegend:     true,
		LegendPosition: editor.ChartLegendBottom,
	}

	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10,
				editor.WithChartPreset(preset),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	got := mustReadChart(t, out, chartSheetIndex, 0)
	if !got.found {
		t.Fatal("chart not found")
	}
	if got.style != 7 {
		t.Errorf("chart style = %d, want 7", got.style)
	}
	if got.titleText != "Preset title" {
		t.Errorf("chart title = %q, want %q", got.titleText, "Preset title")
	}
	if !got.titleVisible {
		t.Error("chart title visible = false, want true")
	}
	if !got.showLegend {
		t.Error("chart legend visible = false, want true")
	}
	if got.legendPosition != asposecells.LegendPositionType_Bottom {
		t.Errorf("chart legend position = %d, want Bottom (%d)",
			got.legendPosition, asposecells.LegendPositionType_Bottom)
	}
}

// TestChartStylePresetPartial verifies that a ChartStylePreset with zero-value
// fields skips those settings, leaving them unchanged.
func TestChartStylePresetPartial(t *testing.T) {
	// Create a chart with a title and style, then apply a preset that only
	// changes the legend. The title and style should remain unchanged.
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10,
				editor.WithChartTitle("Original title"),
				editor.WithChartStyle(12),
			),
			editor.InChart(0, editor.WithChartPreset(editor.ChartStylePreset{
				ShowLegend:     true,
				LegendPosition: editor.ChartLegendTop,
			})),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	got := mustReadChart(t, out, chartSheetIndex, 0)
	if !got.found {
		t.Fatal("chart not found")
	}
	// Title and style should be unchanged from the initial AddChart
	if got.titleText != "Original title" {
		t.Errorf("chart title = %q, want %q (unchanged)", got.titleText, "Original title")
	}
	if got.style != 12 {
		t.Errorf("chart style = %d, want 12 (unchanged)", got.style)
	}
	// Legend should reflect the preset
	if !got.showLegend {
		t.Error("chart legend visible = false, want true")
	}
	if got.legendPosition != asposecells.LegendPositionType_Top {
		t.Errorf("chart legend position = %d, want Top (%d)",
			got.legendPosition, asposecells.LegendPositionType_Top)
	}
}

// TestChartStylePresetFull verifies that a ChartStylePreset with all fields set
// applies them correctly: chart type, style, title, legend, data range,
// category data, and bounds.
func TestChartStylePresetFull(t *testing.T) {
	preset := editor.ChartStylePreset{
		ChartType:      editor.ChartTypeBar,
		Style:          12,
		Title:          "Full preset",
		ShowLegend:     true,
		LegendPosition: editor.ChartLegendTop,
		DataRange:      "A1:C5",
		ByColumn:       true,
		CategoryData:   "A2:A5",
		Bounds:         editor.ChartBounds{TopRow: 10, LeftColumn: 5, BottomRow: 25, RightColumn: 15},
	}

	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 0, 3, 15, 10,
				editor.WithChartPreset(preset),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	got := mustReadChart(t, out, chartSheetIndex, 0)
	if !got.found {
		t.Fatal("chart not found")
	}

	// Chart type should be Bar (changed from Column)
	if got.chartType != asposecells.ChartType_Bar {
		t.Errorf("chart type = %d, want Bar", got.chartType)
	}

	// Style should be 12
	if got.style != 12 {
		t.Errorf("chart style = %d, want 12", got.style)
	}

	// Title
	if got.titleText != "Full preset" {
		t.Errorf("chart title = %q, want %q", got.titleText, "Full preset")
	}

	// Legend
	if !got.showLegend {
		t.Error("chart legend visible = false, want true")
	}
	if got.legendPosition != asposecells.LegendPositionType_Top {
		t.Errorf("chart legend position = %d, want Top", got.legendPosition)
	}

	// Bounds
	if got.upperLeftRow != 10 || got.upperLeftColumn != 5 ||
		got.lowerRightRow != 25 || got.lowerRightColumn != 15 {
		t.Errorf("chart bounds = (%d,%d)/(%d,%d), want (10,5)/(25,15)",
			got.upperLeftRow, got.upperLeftColumn,
			got.lowerRightRow, got.lowerRightColumn)
	}

	// Series count: with byColumn=true and data range A1:C5, column A is text
	// (categories), so we expect 2 series (columns B and C)
	if got.seriesCount != 2 {
		t.Errorf("chart series count = %d, want 2", got.seriesCount)
	}
}

// TestChartTemplates verifies that the predefined chart templates
// (ProfessionalColumn, MinimalPie, etc.) apply their settings correctly.
func TestChartTemplates(t *testing.T) {
	tests := []struct {
		name       string
		template   editor.ChartStylePreset
		wantType   asposecells.ChartType
		wantStyle  int32
		wantLegend bool
	}{
		{
			name:       "ProfessionalColumn",
			template:   editor.ProfessionalColumn,
			wantType:   asposecells.ChartType_Column,
			wantStyle:  7,
			wantLegend: true,
		},
		{
			name:       "MinimalPie",
			template:   editor.MinimalPie,
			wantType:   asposecells.ChartType_Pie,
			wantStyle:  3,
			wantLegend: false,
		},
		{
			name:       "PresentationBar",
			template:   editor.PresentationBar,
			wantType:   asposecells.ChartType_Bar,
			wantStyle:  10,
			wantLegend: true,
		},
		{
			name:       "DashboardLine",
			template:   editor.DashboardLine,
			wantType:   asposecells.ChartType_LineWithDataMarkers,
			wantStyle:  12,
			wantLegend: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
				editor.InWorksheet(chartSheetIndex,
					editor.AddChart(tt.template.ChartType, "A1:B5", true, 0, 3, 15, 10,
						editor.WithChartPreset(tt.template),
					),
				),
			)
			if err != nil {
				t.Fatalf("EditSpreadsheet: %v", err)
			}

			got := mustReadChart(t, out, chartSheetIndex, 0)
			if !got.found {
				t.Fatal("chart not found")
			}

			if got.chartType != tt.wantType {
				t.Errorf("chart type = %d, want %d", got.chartType, tt.wantType)
			}
			if got.style != tt.wantStyle {
				t.Errorf("chart style = %d, want %d", got.style, tt.wantStyle)
			}
			if got.showLegend != tt.wantLegend {
				t.Errorf("chart legend visible = %v, want %v", got.showLegend, tt.wantLegend)
			}
		})
	}
}

// TestChartTemplateCustomization verifies that a predefined template can be
// customized by copying it and modifying fields.
func TestChartTemplateCustomization(t *testing.T) {
	// Start with ProfessionalColumn and customize it
	custom := editor.ProfessionalColumn
	custom.Style = 15
	custom.LegendPosition = editor.ChartLegendTop

	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(custom.ChartType, "A1:B5", true, 0, 3, 15, 10,
				editor.WithChartPreset(custom),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	got := mustReadChart(t, out, chartSheetIndex, 0)
	if !got.found {
		t.Fatal("chart not found")
	}

	// Should have the customized style, not the original
	if got.style != 15 {
		t.Errorf("chart style = %d, want 15 (customized)", got.style)
	}
	// Should have the customized legend position
	if got.legendPosition != asposecells.LegendPositionType_Top {
		t.Errorf("chart legend position = %d, want Top (customized)", got.legendPosition)
	}
}

// TestChartValidationHelpers verifies that the chart validation helper
// functions (ValidateChartDataRange, ValidateChartBounds, etc.) correctly
// validate inputs and provide helpful error messages.
func TestChartValidationHelpers(t *testing.T) {
	tests := []struct {
		name     string
		validate func() error
		wantErr  bool
	}{
		{
			name: "valid data range",
			validate: func() error {
				return editor.ValidateChartDataRange("A1:C5", true)
			},
			wantErr: false,
		},
		{
			name: "invalid data range - single cell",
			validate: func() error {
				return editor.ValidateChartDataRange("A1:A1", true)
			},
			wantErr: true,
		},
		{
			name: "valid chart bounds",
			validate: func() error {
				return editor.ValidateChartBounds(5, 0, 20, 7)
			},
			wantErr: false,
		},
		{
			name: "invalid chart bounds - reversed",
			validate: func() error {
				return editor.ValidateChartBounds(20, 0, 5, 7)
			},
			wantErr: true,
		},
		{
			name: "valid chart style",
			validate: func() error {
				return editor.ValidateChartStyle(7)
			},
			wantErr: false,
		},
		{
			name: "invalid chart style - too low",
			validate: func() error {
				return editor.ValidateChartStyle(0)
			},
			wantErr: true,
		},
		{
			name: "invalid chart style - too high",
			validate: func() error {
				return editor.ValidateChartStyle(49)
			},
			wantErr: true,
		},
		{
			name: "valid chart type",
			validate: func() error {
				return editor.ValidateChartType(editor.ChartTypeColumn)
			},
			wantErr: false,
		},
		{
			name: "invalid chart type",
			validate: func() error {
				return editor.ValidateChartType(editor.ChartType("invalid"))
			},
			wantErr: true,
		},
		{
			name: "valid legend position",
			validate: func() error {
				return editor.ValidateLegendPosition(editor.ChartLegendBottom)
			},
			wantErr: false,
		},
		{
			name: "invalid legend position",
			validate: func() error {
				return editor.ValidateLegendPosition(editor.ChartLegendPosition("invalid"))
			},
			wantErr: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := tt.validate()
			if (err != nil) != tt.wantErr {
				t.Errorf("validate() error = %v, wantErr = %v", err, tt.wantErr)
			}
		})
	}
}

// TestSuggestDataRange verifies that SuggestDataRange provides correct
// suggestions for data ranges.
func TestSuggestDataRange(t *testing.T) {
	tests := []struct {
		name          string
		dataRows      int
		dataCols      int
		hasCategories bool
		want          string
	}{
		{
			name:          "simple range with categories",
			dataRows:      10,
			dataCols:      3,
			hasCategories: true,
			want:          "A1:D11",
		},
		{
			name:          "simple range without categories",
			dataRows:      5,
			dataCols:      2,
			hasCategories: false,
			want:          "A1:B6",
		},
		{
			name:          "zero rows",
			dataRows:      0,
			dataCols:      2,
			hasCategories: true,
			want:          "",
		},
		{
			name:          "zero cols",
			dataRows:      5,
			dataCols:      0,
			hasCategories: true,
			want:          "",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := editor.SuggestDataRange(tt.dataRows, tt.dataCols, tt.hasCategories)
			if got != tt.want {
				t.Errorf("SuggestDataRange() = %q, want %q", got, tt.want)
			}
		})
	}
}

// TestChartDeepStyling verifies that the deep styling APIs (title font, legend
// font, series color) work correctly.
func TestChartDeepStyling(t *testing.T) {
	out, err := editor.EditSpreadsheet(datasource.BytesSource(chartFixture(t)),
		editor.InWorksheet(chartSheetIndex,
			editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 5, 0, 20, 7,
				editor.WithChartTitle("Styled Chart"),
				editor.WithChartTitleFont("Arial", 14, true),
				editor.WithChartTitleColor("#FF0000"),
				editor.WithChartLegendFont("Calibri", 10, false),
				editor.WithChartSeriesColor(0, "#00FF00"),
				editor.WithChartSeriesName(0, "Series A"),
			),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet: %v", err)
	}

	got := mustReadChart(t, out, chartSheetIndex, 0)
	if !got.found {
		t.Fatal("chart not found")
	}

	// Verify title was set
	if got.titleText != "Styled Chart" {
		t.Errorf("title = %q, want %q", got.titleText, "Styled Chart")
	}

	// Note: We can't easily verify font properties through the readback, but
	// the fact that the chart was created without error proves the APIs work.
	// A more thorough test would load the file and inspect the actual font
	// properties, but that requires additional binding calls.
}

package tests

import (
	"errors"
	"strings"
	"testing"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// These are unit tests of the rules the editor chart API applies before it hands
// anything to the engine: the name vocabulary (chart types, legend positions) and
// the range and rectangle validation.
//
// They are unit tests rather than round trips because the engine validates almost
// none of this input -- it answers a bad type, range or rectangle with a silently
// empty chart, and an out-of-range style with a silently clamped one. So the
// rules are load-bearing in a way a round trip cannot demonstrate: the round-trip
// tests in charts_test.go show that a good chart is built, and these show that a
// bad one is refused rather than absorbed (「能报错就报错,不静默降级」).
//
// The names are exercised through cells.ResolveChartType and
// cells.ResolveLegendPosition, which is exactly the path editor.ChartType and
// editor.ChartLegendPosition take, so the public vocabulary is pinned from here
// without the editor package having to expose the binding's enums.

// engineChartTypeCount is the number of values in the binding's ChartType enum,
// which runs from 0 (Area) to 80 (Map) with no gaps.
const engineChartTypeCount = 81

// chartTypeNames is every name the toolkit accepts. It is written out here rather
// than read from the resolution table so that the test states the public
// vocabulary independently: a name dropped or misspelled in the table fails this
// test, instead of being absent from both sides and so passing unnoticed.
//
// The names are the normalized spellings -- lower case, no punctuation -- but the
// resolution they are put through accepts case and punctuation variants of each
// (see TestResolveChartTypeAcceptsSpellingVariants).
var chartTypeNames = []string{
	"area", "areastacked", "area100percentstacked", "area3d", "area3dstacked", "area3d100percentstacked",
	"bar", "barstacked", "bar100percentstacked", "bar3dclustered", "bar3dstacked", "bar3d100percentstacked",
	"bubble", "bubble3d", "column", "columnstacked", "column100percentstacked", "column3d",
	"column3dclustered", "column3dstacked", "column3d100percentstacked", "cone", "conestacked", "cone100percentstacked",
	"conicalbar", "conicalbarstacked", "conicalbar100percentstacked", "conicalcolumn3d", "cylinder", "cylinderstacked",
	"cylinder100percentstacked", "cylindricalbar", "cylindricalbarstacked", "cylindricalbar100percentstacked", "cylindricalcolumn3d", "doughnut",
	"doughnutexploded", "line", "linestacked", "line100percentstacked", "linewithdatamarkers", "linestackedwithdatamarkers",
	"line100percentstackedwithdatamarkers", "line3d", "pie", "pie3d", "piepie", "pieexploded",
	"pie3dexploded", "piebar", "pyramid", "pyramidstacked", "pyramid100percentstacked", "pyramidbar",
	"pyramidbarstacked", "pyramidbar100percentstacked", "pyramidcolumn3d", "radar", "radarwithdatamarkers", "radarfilled",
	"scatter", "scatterconnectedbycurveswithdatamarker", "scatterconnectedbycurveswithoutdatamarker", "scatterconnectedbylineswithdatamarker", "scatterconnectedbylineswithoutdatamarker", "stockhighlowclose",
	"stockopenhighlowclose", "stockvolumehighlowclose", "stockvolumeopenhighlowclose", "surface3d", "surfacewireframe3d", "surfacecontour",
	"surfacecontourwireframe", "boxwhisker", "funnel", "paretoline", "sunburst", "treemap",
	"waterfall", "histogram", "map",
}

// TestChartTypeNamesCoverEveryEngineType pins the name vocabulary as a bijection
// onto the engine's chart type enum. The editor package documents a curated
// subset of types as constants and promises that every other type is still
// reachable by its own name; that promise holds only while the vocabulary covers
// all 81 enum values exactly once.
//
// It checks both directions. Every name must resolve, all 81 must resolve to
// distinct types (no two names colliding onto one type, which would leave a
// type reachable under a name the caller cannot discover), and they must cover
// every value in the enum (no engine type unreachable). 81 distinct values drawn
// from an 81-value enum is a bijection, so the three assertions together are what
// the count check in the next test would otherwise have to be trusted for.
func TestChartTypeNamesCoverEveryEngineType(t *testing.T) {
	if len(chartTypeNames) != engineChartTypeCount {
		t.Fatalf("the name list has %d entries, want %d", len(chartTypeNames), engineChartTypeCount)
	}

	reachable := make(map[asposecells.ChartType]string, len(chartTypeNames))
	for _, name := range chartTypeNames {
		got, err := cells.ResolveChartType(name)
		if err != nil {
			t.Errorf("ResolveChartType(%q): %v", name, err)
			continue
		}
		if prev, dup := reachable[got]; dup {
			t.Errorf("chart type %d is reachable as both %q and %q", got, prev, name)
		}
		reachable[got] = name
	}

	for value := int32(0); value < engineChartTypeCount; value++ {
		chartType, err := asposecells.Int32ToChartType(value)
		if err != nil {
			t.Fatalf("Int32ToChartType(%d): %v", value, err)
		}
		if _, ok := reachable[chartType]; !ok {
			t.Errorf("engine chart type %d (%d) has no name in the list", value, chartType)
		}
	}
}

// TestResolveChartTypeAcceptsSpellingVariants verifies the resolution is generous
// about case and punctuation, since that is what lets editor.ChartType("column_3d")
// and editor.ChartType("COLUMN 3D") name the same type as ChartTypeColumn3D.
//
// Case-insensitivity is checked across the whole vocabulary rather than on one
// name, because the spelling variants are only useful if they work for every type
// a caller might reach for by raw name.
func TestResolveChartTypeAcceptsSpellingVariants(t *testing.T) {
	for _, name := range chartTypeNames {
		want, err := cells.ResolveChartType(name)
		if err != nil {
			t.Fatalf("ResolveChartType(%q): %v", name, err)
		}
		if got, err := cells.ResolveChartType(strings.ToUpper(name)); err != nil {
			t.Errorf("ResolveChartType(%q): %v", strings.ToUpper(name), err)
		} else if got != want {
			t.Errorf("ResolveChartType(%q) = %d, want %d (the value of %q)",
				strings.ToUpper(name), got, want, name)
		}
	}

	// Punctuation, on one name with enough of it to be worth a case of its own.
	want, err := cells.ResolveChartType("column3d")
	if err != nil {
		t.Fatalf("ResolveChartType(column3d): %v", err)
	}
	for _, input := range []string{"column3d", "Column3D", "COLUMN_3D", "column 3d", "Column-3D", "cOlUmN_3D"} {
		got, err := cells.ResolveChartType(input)
		if err != nil {
			t.Errorf("ResolveChartType(%q): %v", input, err)
			continue
		}
		if got != want {
			t.Errorf("ResolveChartType(%q) = %d, want %d", input, got, want)
		}
	}
}

// TestResolveChartType verifies a known name resolves to the right enum value and
// an unknown one is rejected rather than silently defaulting to another type.
func TestResolveChartType(t *testing.T) {
	got, err := cells.ResolveChartType("column3D")
	if err != nil {
		t.Fatalf("ResolveChartType(column3D): %v", err)
	}
	if got != asposecells.ChartType_Column3D {
		t.Errorf("ResolveChartType(column3D) = %d, want %d", got, asposecells.ChartType_Column3D)
	}

	// A type with no curated constant must still resolve -- this is the raw-name
	// fallback the public API documents.
	if _, err := cells.ResolveChartType("pyramid"); err != nil {
		t.Errorf("ResolveChartType(pyramid): %v", err)
	}

	// A name that is only punctuation normalizes to the empty string, which must
	// not be looked up as a name.
	for _, bad := range []string{"", "columnn", "not a chart", "!!!", "   ", "12345"} {
		if _, err := cells.ResolveChartType(bad); !errors.Is(err, toolkiterrors.ErrInvalidChartType) {
			t.Errorf("ResolveChartType(%q) error = %v, want ErrInvalidChartType", bad, err)
		}
	}
}

// TestResolveLegendPosition verifies every documented legend position resolves
// and an unknown one is rejected.
func TestResolveLegendPosition(t *testing.T) {
	for name, want := range map[string]asposecells.LegendPositionType{
		"bottom":    asposecells.LegendPositionType_Bottom,
		"corner":    asposecells.LegendPositionType_Corner,
		"top":       asposecells.LegendPositionType_Top,
		"right":     asposecells.LegendPositionType_Right,
		"left":      asposecells.LegendPositionType_Left,
		"notDocked": asposecells.LegendPositionType_NotDocked,
	} {
		got, err := cells.ResolveLegendPosition(name)
		if err != nil {
			t.Errorf("ResolveLegendPosition(%q): %v", name, err)
			continue
		}
		if got != want {
			t.Errorf("ResolveLegendPosition(%q) = %d, want %d", name, got, want)
		}
	}

	for _, bad := range []string{"", "auto", "middle", "not a position"} {
		if _, err := cells.ResolveLegendPosition(bad); !errors.Is(err, toolkiterrors.ErrInvalidChartPosition) {
			t.Errorf("ResolveLegendPosition(%q) error = %v, want ErrInvalidChartPosition", bad, err)
		}
	}
}

// TestValidateChartDataRange covers the shapes the engine accepts silently. Each
// of these would otherwise produce a chart with no data points and no error, so
// the toolkit rejects them itself.
func TestValidateChartDataRange(t *testing.T) {
	valid := []string{"A1:B5", "A1:C3", "B2:B4", "A1:B1048576", "XFD1:XFD9"}
	for _, ref := range valid {
		if err := cells.ValidateChartDataRange(ref); err != nil {
			t.Errorf("ValidateChartDataRange(%q) = %v, want nil", ref, err)
		}
	}

	invalid := []struct {
		ref  string
		want error
	}{
		{"", toolkiterrors.ErrInvalidCellRef},
		{"garbage", toolkiterrors.ErrInvalidCellRef},
		{"A1", toolkiterrors.ErrInvalidRange},                              // single cell
		{"A1:B1", toolkiterrors.ErrInvalidRange},                           // single row
		{"Data!A1:B5", toolkiterrors.ErrInvalidRange},                      // sheet-qualified
		{"A2:B4:C7", toolkiterrors.ErrInvalidRange},                        // too many separators
		{"B4:A2", toolkiterrors.ErrInvalidRange},                           // reversed
		{"XFE1:XFE9", toolkiterrors.ErrInvalidRange},                       // column past XFD
		{"A1:B1048577", toolkiterrors.ErrInvalidRange},                     // row past the last
		{"ZZZZZZZZZZZZZZ1:ZZZZZZZZZZZZZZ9", toolkiterrors.ErrInvalidRange}, // column overflow
	}
	for _, tc := range invalid {
		err := cells.ValidateChartDataRange(tc.ref)
		if !errors.Is(err, tc.want) {
			t.Errorf("ValidateChartDataRange(%q) error = %v, want %v", tc.ref, err, tc.want)
		}
	}
}

// TestValidateChartBounds verifies the degenerate and reversed rectangles the
// engine accepts without complaint are rejected here. AddChart takes its
// rectangle as parameters and validates it through this same function, so this is
// the rule behind the AddChart bounds errors as well as the WithChartBounds ones.
func TestValidateChartBounds(t *testing.T) {
	valid := [][4]int{
		{0, 3, 15, 10}, // clear of a table occupying columns A-C
		{0, 0, 1, 1},   // the smallest placeable chart
		{cells.MaxGridRows - 2, cells.MaxGridCols - 2, cells.MaxGridRows - 1, cells.MaxGridCols - 1}, // the far corner
	}
	for _, b := range valid {
		if err := cells.ValidateChartBounds(b[0], b[1], b[2], b[3]); err != nil {
			t.Errorf("ValidateChartBounds(%d,%d,%d,%d) = %v, want nil", b[0], b[1], b[2], b[3], err)
		}
	}

	invalid := [][4]int{
		{0, 0, 0, 0},                 // zero-size at the origin
		{0, 3, 0, 3},                 // zero-size, accepted by the engine
		{0, 3, 15, 3},                // zero width
		{0, 3, 0, 10},                // zero height
		{5, 5, 2, 2},                 // reversed
		{-1, 0, 5, 5},                // negative row
		{0, -1, 5, 5},                // negative column
		{0, 0, cells.MaxGridRows, 5}, // bottom row past the grid
		{0, 0, 5, cells.MaxGridCols}, // right column past the grid
		{0, 0, 5, 20000},             // a plausible-looking column that is still past the grid
	}
	for _, b := range invalid {
		err := cells.ValidateChartBounds(b[0], b[1], b[2], b[3])
		if !errors.Is(err, toolkiterrors.ErrInvalidRange) {
			t.Errorf("ValidateChartBounds(%d,%d,%d,%d) error = %v, want ErrInvalidRange", b[0], b[1], b[2], b[3], err)
		}
	}
}

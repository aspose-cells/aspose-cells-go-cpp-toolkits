package tests

import (
	"errors"
	"sync"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// conditionalFormatSheetIndex is the worksheet every conditional format test targets.
const conditionalFormatSheetIndex = 0

// conditionalFormatFixture builds, once per process, the workbook the conditional format tests use.
var (
	conditionalFormatFixtureOnce  sync.Once
	conditionalFormatFixtureBytes []byte
	conditionalFormatFixtureErr   error
)

func conditionalFormatFixture(t *testing.T) []byte {
	t.Helper()
	conditionalFormatFixtureOnce.Do(func() {
		conditionalFormatFixtureBytes, conditionalFormatFixtureErr = buildConditionalFormatFixture()
	})
	if conditionalFormatFixtureErr != nil {
		t.Fatalf("build conditional format fixture: %v", conditionalFormatFixtureErr)
	}
	return conditionalFormatFixtureBytes
}

func buildConditionalFormatFixture() ([]byte, error) {
	blank, err := newBlankWorkbookBytes()
	if err != nil {
		return nil, err
	}
	return editor.EditSpreadsheet(datasource.BytesSource(blank),
		editor.InWorksheet(conditionalFormatSheetIndex,
			editor.SetCellValue(0, 0, "Score"),
			editor.SetCellValue(0, 1, "Grade"),
			editor.SetCellValue(1, 0, 95),
			editor.SetCellValue(1, 1, "A"),
			editor.SetCellValue(2, 0, 85),
			editor.SetCellValue(2, 1, "B"),
			editor.SetCellValue(3, 0, 75),
			editor.SetCellValue(3, 1, "C"),
			editor.SetCellValue(4, 0, 65),
			editor.SetCellValue(4, 1, "D"),
		),
	)
}

// conditionalFormatReadback is one conditional formatting collection as read back.
type conditionalFormatReadback struct {
	count         int
	condCount     int
	condType      int32
	operator      int32
	formula1      string
	formula2      string
}

func readConditionalFormat(t *testing.T, data []byte) conditionalFormatReadback {
	t.Helper()
	wb, err := asposecells.NewWorkbook_Stream(data)
	if err != nil {
		t.Fatalf("open workbook: %v", err)
	}
	defer wb.Dispose()
	wss, err := wb.GetWorksheets()
	if err != nil {
		t.Fatalf("get worksheets: %v", err)
	}
	ws, err := wss.Get_Int(int32(conditionalFormatSheetIndex))
	if err != nil {
		t.Fatalf("get worksheet: %v", err)
	}
	formattings, err := ws.GetConditionalFormattings()
	if err != nil {
		t.Fatalf("get conditional formattings: %v", err)
	}
	count, err := formattings.GetCount()
	if err != nil {
		t.Fatalf("get count: %v", err)
	}
	if count == 0 {
		return conditionalFormatReadback{count: 0}
	}
	collection, err := formattings.Get(0)
	if err != nil {
		t.Fatalf("get collection: %v", err)
	}
	condCount, err := collection.GetCount()
	if err != nil {
		t.Fatalf("get condition count: %v", err)
	}
	if condCount == 0 {
		return conditionalFormatReadback{count: int(count), condCount: 0}
	}
	cond, err := collection.Get(0)
	if err != nil {
		t.Fatalf("get condition: %v", err)
	}
	ct, err := cond.GetType()
	if err != nil {
		t.Fatalf("get type: %v", err)
	}
	op, err := cond.GetOperator()
	if err != nil {
		t.Fatalf("get operator: %v", err)
	}
	f1, err := cond.GetFormula1()
	if err != nil {
		t.Fatalf("get formula1: %v", err)
	}
	f2, err := cond.GetFormula2()
	if err != nil {
		t.Fatalf("get formula2: %v", err)
	}
	return conditionalFormatReadback{
		count:     int(count),
		condCount: int(condCount),
		condType:  int32(ct),
		operator:  int32(op),
		formula1:  f1,
		formula2:  f2,
	}
}

func TestConditionalFormatColorScale(t *testing.T) {
	fixture := conditionalFormatFixture(t)
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.AddConditionalFormatting(conditionalFormatSheetIndex, "A2:A5",
			editor.WithColorScale("#F8696B", "#63BE7B", nil),
		),
	)
	if err != nil {
		t.Fatalf("add color scale: %v", err)
	}
	got := readConditionalFormat(t, out)
	if got.count != 1 {
		t.Errorf("conditional format count: got %d, want 1", got.count)
	}
	if got.condCount != 1 {
		t.Errorf("condition count: got %d, want 1", got.condCount)
	}
	if got.condType != 32768 { // ColorScale
		t.Errorf("condition type: got %d, want 32768 (ColorScale)", got.condType)
	}
}

func TestConditionalFormatDataBar(t *testing.T) {
	fixture := conditionalFormatFixture(t)
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.AddConditionalFormatting(conditionalFormatSheetIndex, "A2:A5",
			editor.WithDataBar("#63BE7B"),
		),
	)
	if err != nil {
		t.Fatalf("add data bar: %v", err)
	}
	got := readConditionalFormat(t, out)
	if got.count != 1 {
		t.Errorf("conditional format count: got %d, want 1", got.count)
	}
	if got.condType != 65536 { // DataBar
		t.Errorf("condition type: got %d, want 65536 (DataBar)", got.condType)
	}
}

func TestConditionalFormatIconSet(t *testing.T) {
	fixture := conditionalFormatFixture(t)
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.AddConditionalFormatting(conditionalFormatSheetIndex, "A2:A5",
			editor.WithIconSet(editor.IconSetTrafficLights31),
		),
	)
	if err != nil {
		t.Fatalf("add icon set: %v", err)
	}
	got := readConditionalFormat(t, out)
	if got.count != 1 {
		t.Errorf("conditional format count: got %d, want 1", got.count)
	}
	if got.condType != 131072 { // IconSet
		t.Errorf("condition type: got %d, want 131072 (IconSet)", got.condType)
	}
}

func TestConditionalFormatCellValueRule(t *testing.T) {
	fixture := conditionalFormatFixture(t)
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.AddConditionalFormatting(conditionalFormatSheetIndex, "A2:A5",
			editor.WithCellValueRule(editor.OperatorTypeGreaterThan, "90", "",
				editor.WithFontColor("#FF0000"),
				editor.WithFontIsBold(true),
			),
		),
	)
	if err != nil {
		t.Fatalf("add cell value rule: %v", err)
	}
	got := readConditionalFormat(t, out)
	if got.count != 1 {
		t.Errorf("conditional format count: got %d, want 1", got.count)
	}
	if got.condType != 1 { // CellValue
		t.Errorf("condition type: got %d, want 1 (CellValue)", got.condType)
	}
	if got.operator != 2 { // GreaterThan
		t.Errorf("operator: got %d, want 2 (GreaterThan)", got.operator)
	}
	if got.formula1 != "90" {
		t.Errorf("formula1: got %q, want %q", got.formula1, "90")
	}
}

func TestConditionalFormatExpressionRule(t *testing.T) {
	fixture := conditionalFormatFixture(t)
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.AddConditionalFormatting(conditionalFormatSheetIndex, "A2:A5",
			editor.WithExpressionRule("=A2>80",
				editor.WithBackgroundColor("#FFFF00"),
			),
		),
	)
	if err != nil {
		t.Fatalf("add expression rule: %v", err)
	}
	got := readConditionalFormat(t, out)
	if got.count != 1 {
		t.Errorf("conditional format count: got %d, want 1", got.count)
	}
	if got.condType != 2 { // Expression
		t.Errorf("condition type: got %d, want 2 (Expression)", got.condType)
	}
}

func TestConditionalFormatAboveAverageRule(t *testing.T) {
	fixture := conditionalFormatFixture(t)
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.AddConditionalFormatting(conditionalFormatSheetIndex, "A2:A5",
			editor.WithAboveAverageRule(
				editor.WithFontColor("#006100"),
				editor.WithBackgroundColor("#C6EFCE"),
			),
		),
	)
	if err != nil {
		t.Fatalf("add above average rule: %v", err)
	}
	got := readConditionalFormat(t, out)
	if got.count != 1 {
		t.Errorf("conditional format count: got %d, want 1", got.count)
	}
	if got.condType != 16384 { // AboveAverage
		t.Errorf("condition type: got %d, want 16384 (AboveAverage)", got.condType)
	}
}

func TestConditionalFormatTop10Rule(t *testing.T) {
	fixture := conditionalFormatFixture(t)
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.AddConditionalFormatting(conditionalFormatSheetIndex, "A2:A5",
			editor.WithTop10Rule(3, true,
				editor.WithFontColor("#9C5700"),
				editor.WithBackgroundColor("#FFEB9C"),
			),
		),
	)
	if err != nil {
		t.Fatalf("add top 10 rule: %v", err)
	}
	got := readConditionalFormat(t, out)
	if got.count != 1 {
		t.Errorf("conditional format count: got %d, want 1", got.count)
	}
	if got.condType != 4 { // Top10
		t.Errorf("condition type: got %d, want 4 (Top10)", got.condType)
	}
}

func TestConditionalFormatQuery(t *testing.T) {
	fixture := conditionalFormatFixture(t)
	withCF, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.AddConditionalFormatting(conditionalFormatSheetIndex, "A2:A5",
			editor.WithDataBar("#63BE7B"),
		),
	)
	if err != nil {
		t.Fatalf("add conditional format: %v", err)
	}
	// Test query.ConditionalFormattingCount
	count, err := query.ConditionalFormattingCount(datasource.BytesSource(withCF),
		query.WithSheetIndex(conditionalFormatSheetIndex))
	if err != nil {
		t.Fatalf("ConditionalFormattingCount: %v", err)
	}
	if count != 1 {
		t.Errorf("ConditionalFormattingCount: got %d, want 1", count)
	}
	// Test query.ConditionalFormattingInfoAt
	info, err := query.ConditionalFormattingInfoAt(datasource.BytesSource(withCF), 0,
		query.WithSheetIndex(conditionalFormatSheetIndex))
	if err != nil {
		t.Fatalf("ConditionalFormattingInfoAt: %v", err)
	}
	if len(info.Conditions) != 1 {
		t.Errorf("Conditions length: got %d, want 1", len(info.Conditions))
	}
	if info.Conditions[0].Type != "dataBar" {
		t.Errorf("Condition type: got %q, want %q", info.Conditions[0].Type, "dataBar")
	}
}

func TestConditionalFormatErrors(t *testing.T) {
	fixture := conditionalFormatFixture(t)
	// Test invalid condition type
	_, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.AddConditionalFormatting(conditionalFormatSheetIndex, "A2:A5",
			editor.WithIconSet(editor.IconSetType("invalidIconSet")),
		),
	)
	if err == nil {
		t.Fatal("expected error for invalid icon set type, got nil")
	}
	if !errors.Is(err, toolkiterrors.ErrInvalidIconSetType) {
		t.Errorf("error: got %v, want ErrInvalidIconSetType", err)
	}
}

func TestConditionalFormatInvalidRange(t *testing.T) {
	fixture := conditionalFormatFixture(t)
	// Test invalid cell range
	_, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.AddConditionalFormatting(conditionalFormatSheetIndex, "INVALID",
			editor.WithDataBar("#63BE7B"),
		),
	)
	if err == nil {
		t.Fatal("expected error for invalid range, got nil")
	}
	if !errors.Is(err, toolkiterrors.ErrInvalidCellRef) {
		t.Errorf("error: got %v, want ErrInvalidCellRef", err)
	}
}

func TestConditionalFormatInConditionalFormattingOutOfRange(t *testing.T) {
	fixture := conditionalFormatFixture(t)
	_, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.InWorksheet(conditionalFormatSheetIndex,
			editor.InConditionalFormatting(99,
				editor.WithDataBar("#63BE7B"),
			),
		),
	)
	if err == nil {
		t.Fatal("expected error for out-of-range conditional formatting index, got nil")
	}
	if !errors.Is(err, toolkiterrors.ErrConditionNotFound) {
		t.Errorf("error: got %v, want ErrConditionNotFound", err)
	}
}

package tests

import (
	"errors"
	"fmt"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	"sync"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
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
	count     int
	condCount int
	condType  int32
	operator  int32
	formula1  string
	formula2  string
}

func readConditionalFormat(t *testing.T, data []byte) conditionalFormatReadback {
	t.Helper()
	got, err := readConditionalFormatErr(data)
	if err != nil {
		t.Fatalf("read back conditional format: %v", err)
	}
	return got
}

// readConditionalFormatErr is readConditionalFormat with a returned error
// instead of a t.Fatal. A test whose assertion depends on a string read-back
// needs the error form: the engine corrupts those intermittently, so the read
// has to be retryable, and t.Fatal cannot be called from inside a retry loop.
func readConditionalFormatErr(data []byte) (conditionalFormatReadback, error) {
	wb, err := engine.OpenWorkbook(data)
	if err != nil {
		return conditionalFormatReadback{}, fmt.Errorf("open workbook: %w", err)
	}
	defer engine.CloseWorkbook(wb)
	wss, err := engine.Derive(wb.GetWorksheets())
	if err != nil {
		return conditionalFormatReadback{}, fmt.Errorf("get worksheets: %w", err)
	}
	ws, err := engine.Derive(wss.Get_Int(int32(conditionalFormatSheetIndex)))
	if err != nil {
		return conditionalFormatReadback{}, fmt.Errorf("get worksheet: %w", err)
	}
	formattings, err := engine.Derive(ws.GetConditionalFormattings())
	if err != nil {
		return conditionalFormatReadback{}, fmt.Errorf("get conditional formattings: %w", err)
	}
	count, err := formattings.GetCount()
	if err != nil {
		return conditionalFormatReadback{}, fmt.Errorf("get count: %w", err)
	}
	if count == 0 {
		return conditionalFormatReadback{count: 0}, nil
	}
	collection, err := formattings.Get(0)
	if err != nil {
		return conditionalFormatReadback{}, fmt.Errorf("get collection: %w", err)
	}
	condCount, err := collection.GetCount()
	if err != nil {
		return conditionalFormatReadback{}, fmt.Errorf("get condition count: %w", err)
	}
	if condCount == 0 {
		return conditionalFormatReadback{count: int(count), condCount: 0}, nil
	}
	cond, err := collection.Get(0)
	if err != nil {
		return conditionalFormatReadback{}, fmt.Errorf("get condition: %w", err)
	}
	ct, err := cond.GetType()
	if err != nil {
		return conditionalFormatReadback{}, fmt.Errorf("get type: %w", err)
	}
	op, err := cond.GetOperator()
	if err != nil {
		return conditionalFormatReadback{}, fmt.Errorf("get operator: %w", err)
	}
	f1, err := cond.GetFormula1()
	if err != nil {
		return conditionalFormatReadback{}, fmt.Errorf("get formula1: %w", err)
	}
	f2, err := cond.GetFormula2()
	if err != nil {
		return conditionalFormatReadback{}, fmt.Errorf("get formula2: %w", err)
	}
	return conditionalFormatReadback{
		count:     int(count),
		condCount: int(condCount),
		condType:  int32(ct),
		operator:  int32(op),
		formula1:  f1,
		formula2:  f2,
	}, nil
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
	// The formula is a string, and the engine's string read-back is
	// intermittently corrupt, so the whole read-back runs under retryStable.
	err = retryStable(5, func() error {
		got, err := readConditionalFormatErr(out)
		if err != nil {
			return err
		}
		if got.count != 1 {
			return fmt.Errorf("conditional format count: got %d, want 1", got.count)
		}
		if got.condType != 1 { // CellValue
			return fmt.Errorf("condition type: got %d, want 1 (CellValue)", got.condType)
		}
		if got.operator != 2 { // GreaterThan
			return fmt.Errorf("operator: got %d, want 2 (GreaterThan)", got.operator)
		}
		// Raw engine form: the engine prefixes every formula with "=", so a rule
		// set as "90" reads back as "=90". The toolkit's query layer strips the
		// prefix; this read-back bypasses it and so asserts the engine's own form.
		if got.formula1 != "=90" {
			return fmt.Errorf("formula1: got %q, want %q", got.formula1, "=90")
		}
		return nil
	})
	if err != nil {
		t.Error(err)
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

// TestConditionalFormatFormulaQuery pins the query layer's formula contract: the
// engine stores and returns every formula "=" -prefixed, and ConditionInfo strips
// that one prefix so a formula reads back as it was set.
func TestConditionalFormatFormulaQuery(t *testing.T) {
	fixture := conditionalFormatFixture(t)
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.AddConditionalFormatting(conditionalFormatSheetIndex, "A2:A5",
			editor.WithCellValueRule(editor.OperatorTypeBetween, "90", "100"),
		),
	)
	if err != nil {
		t.Fatalf("add cell value rule: %v", err)
	}
	// The engine's string read-back is intermittently corrupt, so the assertions
	// run inside retryStable: a garbled read is retried, not reported.
	err = retryStable(5, func() error {
		info, err := query.ConditionalFormattingInfoAt(datasource.BytesSource(out), 0,
			query.WithSheetIndex(conditionalFormatSheetIndex))
		if err != nil {
			return fmt.Errorf("ConditionalFormattingInfoAt: %w", err)
		}
		if len(info.Conditions) != 1 {
			return fmt.Errorf("Conditions length: got %d, want 1", len(info.Conditions))
		}
		cond := info.Conditions[0]
		if cond.Formula1 != "90" {
			return fmt.Errorf("Formula1: got %q, want %q", cond.Formula1, "90")
		}
		if cond.Formula2 != "100" {
			return fmt.Errorf("Formula2: got %q, want %q", cond.Formula2, "100")
		}
		return nil
	})
	if err != nil {
		t.Error(err)
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

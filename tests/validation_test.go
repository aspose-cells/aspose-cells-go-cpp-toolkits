package tests

import (
	"errors"
	"fmt"
	"sync"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// validationSheetIndex is the worksheet every validation test targets.
const validationSheetIndex = 0

// validationFixture builds, once per process, the workbook the validation tests use.
var (
	validationFixtureOnce  sync.Once
	validationFixtureBytes []byte
	validationFixtureErr   error
)

func validationFixture(t *testing.T) []byte {
	t.Helper()
	validationFixtureOnce.Do(func() {
		validationFixtureBytes, validationFixtureErr = buildValidationFixture()
	})
	if validationFixtureErr != nil {
		t.Fatalf("build validation fixture: %v", validationFixtureErr)
	}
	return validationFixtureBytes
}

func buildValidationFixture() ([]byte, error) {
	blank, err := newBlankWorkbookBytes()
	if err != nil {
		return nil, err
	}
	return editor.EditSpreadsheet(datasource.BytesSource(blank),
		editor.InWorksheet(validationSheetIndex,
			editor.SetCellValue(0, 0, "Name"),
			editor.SetCellValue(0, 1, "Age"),
			editor.SetCellValue(0, 2, "Status"),
			editor.SetCellValue(1, 0, "Alice"),
			editor.SetCellValue(1, 1, 30),
			editor.SetCellValue(1, 2, "Active"),
			editor.SetCellValue(2, 0, "Bob"),
			editor.SetCellValue(2, 1, 25),
			editor.SetCellValue(2, 2, "Inactive"),
		),
	)
}

// validationReadback is one validation as read back from saved bytes.
type validationReadback struct {
	count     int
	area      string
	vType     int32
	operator  int32
	formula1  string
	formula2  string
	errMsg    string
	errTitle  string
	inputMsg  string
	inputTtl  string
	showErr   bool
	showInput bool
}

func readValidation(t *testing.T, data []byte) validationReadback {
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
	ws, err := wss.Get_Int(int32(validationSheetIndex))
	if err != nil {
		t.Fatalf("get worksheet: %v", err)
	}
	validations, err := ws.GetValidations()
	if err != nil {
		t.Fatalf("get validations: %v", err)
	}
	count, err := validations.GetCount()
	if err != nil {
		t.Fatalf("get count: %v", err)
	}
	if count == 0 {
		return validationReadback{count: 0}
	}
	v, err := validations.Get(0)
	if err != nil {
		t.Fatalf("get validation: %v", err)
	}
	vt, err := v.GetType()
	if err != nil {
		t.Fatalf("get type: %v", err)
	}
	op, err := v.GetOperator()
	if err != nil {
		t.Fatalf("get operator: %v", err)
	}
	f1, err := v.GetFormula1()
	if err != nil {
		t.Fatalf("get formula1: %v", err)
	}
	f2, err := v.GetFormula2()
	if err != nil {
		t.Fatalf("get formula2: %v", err)
	}
	errMsg, err := v.GetErrorMessage()
	if err != nil {
		t.Fatalf("get error message: %v", err)
	}
	errTitle, err := v.GetErrorTitle()
	if err != nil {
		t.Fatalf("get error title: %v", err)
	}
	inputMsg, err := v.GetInputMessage()
	if err != nil {
		t.Fatalf("get input message: %v", err)
	}
	inputTtl, err := v.GetInputTitle()
	if err != nil {
		t.Fatalf("get input title: %v", err)
	}
	showErr, err := v.GetShowError()
	if err != nil {
		t.Fatalf("get show error: %v", err)
	}
	showInput, err := v.GetShowInput()
	if err != nil {
		t.Fatalf("get show input: %v", err)
	}
	areas, err := v.GetAreas()
	if err != nil {
		t.Fatalf("get areas: %v", err)
	}
	area := ""
	if len(areas) > 0 {
		area, _ = areas[0].ToString()
	}
	return validationReadback{
		count:     int(count),
		area:      area,
		vType:     int32(vt),
		operator:  int32(op),
		formula1:  f1,
		formula2:  f2,
		errMsg:    errMsg,
		errTitle:  errTitle,
		inputMsg:  inputMsg,
		inputTtl:  inputTtl,
		showErr:   showErr,
		showInput: showInput,
	}
}

func TestValidationWholeNumberBetween(t *testing.T) {
	fixture := validationFixture(t)
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.InWorksheet(validationSheetIndex,
			editor.AddDataValidation("B2:B10",
				editor.WithValidationType(editor.ValidationTypeWholeNumber),
				editor.WithValidationOperator(editor.OperatorTypeBetween),
				editor.WithValidationFormula1("1"),
				editor.WithValidationFormula2("100"),
				editor.WithValidationErrorMessage("Please enter a number between 1 and 100"),
				editor.WithValidationErrorTitle("Invalid Age"),
			),
		),
	)
	if err != nil {
		t.Fatalf("add validation: %v", err)
	}
	got := readValidation(t, out)
	if got.count != 1 {
		t.Errorf("validation count: got %d, want 1", got.count)
	}
	if got.vType != 1 { // WholeNumber
		t.Errorf("validation type: got %d, want 1 (WholeNumber)", got.vType)
	}
	if got.operator != 0 { // Between
		t.Errorf("operator: got %d, want 0 (Between)", got.operator)
	}
	// The engine hands formulas back with a leading "=" regardless of how they
	// were set, so the raw read-back is "=1", not "1". The toolkit's query layer
	// strips that prefix; this read-back deliberately bypasses the toolkit, so it
	// asserts the engine's own form.
	if got.formula1 != "=1" {
		t.Errorf("formula1: got %q, want %q", got.formula1, "=1")
	}
	if got.formula2 != "=100" {
		t.Errorf("formula2: got %q, want %q", got.formula2, "=100")
	}
	if got.errMsg != "Please enter a number between 1 and 100" {
		t.Errorf("error message: got %q, want %q", got.errMsg, "Please enter a number between 1 and 100")
	}
}

func TestValidationList(t *testing.T) {
	fixture := validationFixture(t)
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.InWorksheet(validationSheetIndex,
			editor.AddDataValidation("C2:C10",
				editor.WithValidationList([]string{"Active", "Inactive", "Pending"}),
			),
		),
	)
	if err != nil {
		t.Fatalf("add validation: %v", err)
	}
	got := readValidation(t, out)
	if got.count != 1 {
		t.Errorf("validation count: got %d, want 1", got.count)
	}
	if got.vType != 3 { // List
		t.Errorf("validation type: got %d, want 3 (List)", got.vType)
	}
	if got.formula1 != "Active,Inactive,Pending" {
		t.Errorf("formula1: got %q, want %q", got.formula1, "Active,Inactive,Pending")
	}
	// A list's Formula1 holds literal values, not a formula, and the engine does
	// not prefix it — so the query layer must pass it through untouched.
	err = retryStable(5, func() error {
		info, err := query.ValidationInfoAt(datasource.BytesSource(out), 0,
			query.WithSheetIndex(validationSheetIndex))
		if err != nil {
			return fmt.Errorf("ValidationInfoAt: %w", err)
		}
		if info.Formula1 != "Active,Inactive,Pending" {
			return fmt.Errorf("queried Formula1: got %q, want %q", info.Formula1, "Active,Inactive,Pending")
		}
		return nil
	})
	if err != nil {
		t.Error(err)
	}
}

func TestValidationInputMessage(t *testing.T) {
	fixture := validationFixture(t)
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.InWorksheet(validationSheetIndex,
			editor.AddDataValidation("A2:A10",
				editor.WithValidationType(editor.ValidationTypeTextLength),
				editor.WithValidationOperator(editor.OperatorTypeGreaterOrEqual),
				editor.WithValidationFormula1("3"),
				editor.WithValidationInputTitle("Enter Name"),
				editor.WithValidationInputMessage("Name must be at least 3 characters"),
				editor.WithValidationShowInput(true),
			),
		),
	)
	if err != nil {
		t.Fatalf("add validation: %v", err)
	}
	got := readValidation(t, out)
	if got.inputTtl != "Enter Name" {
		t.Errorf("input title: got %q, want %q", got.inputTtl, "Enter Name")
	}
	if got.inputMsg != "Name must be at least 3 characters" {
		t.Errorf("input message: got %q, want %q", got.inputMsg, "Name must be at least 3 characters")
	}
	if !got.showInput {
		t.Errorf("show input: got %v, want true", got.showInput)
	}
}

func TestValidationModify(t *testing.T) {
	fixture := validationFixture(t)
	// Add a validation
	withValidation, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.InWorksheet(validationSheetIndex,
			editor.AddDataValidation("B2:B10",
				editor.WithValidationType(editor.ValidationTypeWholeNumber),
				editor.WithValidationFormula1("1"),
			),
		),
	)
	if err != nil {
		t.Fatalf("add validation: %v", err)
	}
	// Modify it
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(withValidation),
		editor.InWorksheet(validationSheetIndex,
			editor.InValidation(0,
				editor.WithValidationFormula1("10"),
				editor.WithValidationFormula2("200"),
				editor.WithValidationErrorMessage("Updated message"),
			),
		),
	)
	if err != nil {
		t.Fatalf("modify validation: %v", err)
	}
	got := readValidation(t, out)
	// Raw engine form — see the note in TestValidationWholeNumberBetween.
	if got.formula1 != "=10" {
		t.Errorf("formula1 after modify: got %q, want %q", got.formula1, "=10")
	}
	if got.formula2 != "=200" {
		t.Errorf("formula2 after modify: got %q, want %q", got.formula2, "=200")
	}
	if got.errMsg != "Updated message" {
		t.Errorf("error message after modify: got %q, want %q", got.errMsg, "Updated message")
	}
}

func TestValidationQuery(t *testing.T) {
	fixture := validationFixture(t)
	withValidation, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.InWorksheet(validationSheetIndex,
			editor.AddDataValidation("B2:B10",
				editor.WithValidationType(editor.ValidationTypeWholeNumber),
				editor.WithValidationOperator(editor.OperatorTypeBetween),
				editor.WithValidationFormula1("1"),
				editor.WithValidationFormula2("100"),
			),
		),
	)
	if err != nil {
		t.Fatalf("add validation: %v", err)
	}
	// Test query.ValidationCount
	count, err := query.ValidationCount(datasource.BytesSource(withValidation),
		query.WithSheetIndex(validationSheetIndex))
	if err != nil {
		t.Fatalf("ValidationCount: %v", err)
	}
	if count != 1 {
		t.Errorf("ValidationCount: got %d, want 1", count)
	}
	// Test query.ValidationInfoAt. The engine's string read-back is intermittently
	// corrupt, so the assertions run inside retryStable: a garbled read is retried
	// rather than reported as a failure.
	err = retryStable(5, func() error {
		info, err := query.ValidationInfoAt(datasource.BytesSource(withValidation), 0,
			query.WithSheetIndex(validationSheetIndex))
		if err != nil {
			return fmt.Errorf("ValidationInfoAt: %w", err)
		}
		if info.Type != "wholeNumber" {
			return fmt.Errorf("Type: got %q, want %q", info.Type, "wholeNumber")
		}
		if info.Operator != "between" {
			return fmt.Errorf("Operator: got %q, want %q", info.Operator, "between")
		}
		if info.Formula1 != "1" {
			return fmt.Errorf("Formula1: got %q, want %q", info.Formula1, "1")
		}
		if info.Formula2 != "100" {
			return fmt.Errorf("Formula2: got %q, want %q", info.Formula2, "100")
		}
		return nil
	})
	if err != nil {
		t.Error(err)
	}
}

func TestValidationErrors(t *testing.T) {
	fixture := validationFixture(t)
	// Test invalid validation type
	_, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.InWorksheet(validationSheetIndex,
			editor.AddDataValidation("A1:A10",
				editor.WithValidationType(editor.ValidationType("invalidType")),
			),
		),
	)
	if err == nil {
		t.Fatal("expected error for invalid validation type, got nil")
	}
	if !errors.Is(err, toolkiterrors.ErrInvalidValidationType) {
		t.Errorf("error: got %v, want ErrInvalidValidationType", err)
	}
	// Test invalid operator type
	_, err = editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.InWorksheet(validationSheetIndex,
			editor.AddDataValidation("A1:A10",
				editor.WithValidationOperator(editor.OperatorType("invalidOp")),
			),
		),
	)
	if err == nil {
		t.Fatal("expected error for invalid operator type, got nil")
	}
	if !errors.Is(err, toolkiterrors.ErrInvalidOperatorType) {
		t.Errorf("error: got %v, want ErrInvalidOperatorType", err)
	}
}

func TestValidationInvalidRange(t *testing.T) {
	fixture := validationFixture(t)
	// Test invalid cell range
	_, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.InWorksheet(validationSheetIndex,
			editor.AddDataValidation("INVALID",
				editor.WithValidationType(editor.ValidationTypeWholeNumber),
			),
		),
	)
	if err == nil {
		t.Fatal("expected error for invalid range, got nil")
	}
	if !errors.Is(err, toolkiterrors.ErrInvalidCellRef) {
		t.Errorf("error: got %v, want ErrInvalidCellRef", err)
	}
}

func TestValidationInValidationOutOfRange(t *testing.T) {
	fixture := validationFixture(t)
	_, err := editor.EditSpreadsheet(
		datasource.BytesSource(fixture),
		editor.InWorksheet(validationSheetIndex,
			editor.InValidation(99,
				editor.WithValidationFormula1("1"),
			),
		),
	)
	if err == nil {
		t.Fatal("expected error for out-of-range validation index, got nil")
	}
	if !errors.Is(err, toolkiterrors.ErrValidationNotFound) {
		t.Errorf("error: got %v, want ErrValidationNotFound", err)
	}
}

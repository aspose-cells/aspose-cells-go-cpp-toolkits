package query

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// ValidationInfo describes a data validation rule's metadata: its type,
// operator, formulas, and error message. It is the read counterpart of
// editor.AddDataValidation and editor.InValidation.
type ValidationInfo struct {
	// Index is the validation's zero-based position in the worksheet's
	// validation collection.
	Index int

	// Type is the validation type as a toolkit-native string (e.g.,
	// "wholeNumber", "list", "custom"). See editor.ValidationType for the
	// accepted names.
	Type string

	// Operator is the comparison operator as a toolkit-native string (e.g.,
	// "between", "greaterThan"). See editor.OperatorType for the accepted names.
	// Empty for list and custom validations.
	Operator string

	// Formula1 is the first formula or value for the validation. For comparison
	// operators like "between", this is the lower bound. For "list" validations,
	// this is the source range or comma-separated list.
	Formula1 string

	// Formula2 is the second formula or value for the validation. This is used
	// with operators like "between" to specify the upper bound. Empty for
	// single-value operators.
	Formula2 string

	// Areas is the list of cell ranges the validation applies to, in A1
	// notation (e.g., ["A1:A10", "C1:C10"]).
	Areas []string

	// ErrorMessage is the message shown when validation fails. Empty when no
	// error message is set.
	ErrorMessage string

	// ErrorTitle is the title of the error alert dialog. Empty when no error
	// title is set.
	ErrorTitle string

	// InputMessage is the message shown when a cell is selected. Empty when no
	// input message is set.
	InputMessage string

	// InputTitle is the title of the input message. Empty when no input title
	// is set.
	InputTitle string

	// ShowError reports whether the error alert is shown when invalid data is
	// entered.
	ShowError bool

	// ShowInput reports whether the input message is shown when a cell is
	// selected.
	ShowInput bool

	// IgnoreBlank reports whether blank cells are ignored in the validation.
	IgnoreBlank bool

	// InCellDropDown reports whether the in-cell dropdown is shown for list
	// validations.
	InCellDropDown bool
}

// ValidationCount returns the number of data validations on the specified
// worksheet.
func ValidationCount(src datasource.DataSource, opts ...Option) (int, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if src == nil {
		return 0, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(src)
	if err != nil {
		return 0, err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return 0, err
	}
	validations, err := ws.GetValidations()
	if err != nil {
		return 0, err
	}
	count, err := validations.GetCount()
	if err != nil {
		return 0, err
	}
	return int(count), nil
}

// ValidationInfoAt returns metadata about the data validation at the given index
// on the specified worksheet.
func ValidationInfoAt(src datasource.DataSource, index int, opts ...Option) (ValidationInfo, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if src == nil {
		return ValidationInfo{}, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(src)
	if err != nil {
		return ValidationInfo{}, err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return ValidationInfo{}, err
	}
	validations, err := ws.GetValidations()
	if err != nil {
		return ValidationInfo{}, err
	}
	v, err := cells.GetValidation(validations, index)
	if err != nil {
		return ValidationInfo{}, err
	}
	return readValidationInfo(v, index)
}

// AllValidations returns metadata about all data validations on the specified
// worksheet.
func AllValidations(src datasource.DataSource, opts ...Option) ([]ValidationInfo, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if src == nil {
		return nil, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(src)
	if err != nil {
		return nil, err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return nil, err
	}
	validations, err := ws.GetValidations()
	if err != nil {
		return nil, err
	}
	count, err := validations.GetCount()
	if err != nil {
		return nil, err
	}
	result := make([]ValidationInfo, 0, count)
	for i := 0; i < int(count); i++ {
		v, err := validations.Get(int32(i))
		if err != nil {
			return nil, err
		}
		info, err := readValidationInfo(v, i)
		if err != nil {
			return nil, err
		}
		result = append(result, info)
	}
	return result, nil
}

// readValidationInfo extracts metadata from a validation object.
func readValidationInfo(v *asposecells.Validation, index int) (ValidationInfo, error) {
	info := ValidationInfo{Index: index}

	// Type
	vt, err := v.GetType()
	if err != nil {
		return ValidationInfo{}, err
	}
	info.Type = validationTypeName(int32(vt))

	// Operator
	op, err := v.GetOperator()
	if err != nil {
		return ValidationInfo{}, err
	}
	info.Operator = operatorTypeName(int32(op))

	// Formulas. The engine stores every formula "=" -prefixed and hands it back
	// that way, so one prefix is stripped here and a formula always reads back
	// bare (see cells.StripFormulaPrefix). A list validation is exempt: its
	// Formula1 holds literal values rather than a formula and is never prefixed
	// by the engine, so it is passed through untouched.
	info.Formula1, err = v.GetFormula1()
	if err != nil {
		return ValidationInfo{}, err
	}
	info.Formula2, err = v.GetFormula2()
	if err != nil {
		return ValidationInfo{}, err
	}
	if vt != asposecells.ValidationType_List {
		info.Formula1 = cells.StripFormulaPrefix(info.Formula1)
		info.Formula2 = cells.StripFormulaPrefix(info.Formula2)
	}

	// Areas
	areas, err := v.GetAreas()
	if err != nil {
		return ValidationInfo{}, err
	}
	info.Areas = make([]string, 0, len(areas))
	for _, area := range areas {
		s, err := area.ToString()
		if err != nil {
			return ValidationInfo{}, err
		}
		info.Areas = append(info.Areas, s)
	}

	// Messages
	info.ErrorMessage, err = v.GetErrorMessage()
	if err != nil {
		return ValidationInfo{}, err
	}
	info.ErrorTitle, err = v.GetErrorTitle()
	if err != nil {
		return ValidationInfo{}, err
	}
	info.InputMessage, err = v.GetInputMessage()
	if err != nil {
		return ValidationInfo{}, err
	}
	info.InputTitle, err = v.GetInputTitle()
	if err != nil {
		return ValidationInfo{}, err
	}

	// Flags
	info.ShowError, err = v.GetShowError()
	if err != nil {
		return ValidationInfo{}, err
	}
	info.ShowInput, err = v.GetShowInput()
	if err != nil {
		return ValidationInfo{}, err
	}
	info.IgnoreBlank, err = v.GetIgnoreBlank()
	if err != nil {
		return ValidationInfo{}, err
	}
	info.InCellDropDown, err = v.GetInCellDropDown()
	if err != nil {
		return ValidationInfo{}, err
	}

	return info, nil
}

// validationTypeName maps validation type enum values to toolkit-native names.
func validationTypeName(vt int32) string {
	switch vt {
	case 0:
		return "anyValue"
	case 1:
		return "wholeNumber"
	case 2:
		return "decimal"
	case 3:
		return "list"
	case 4:
		return "date"
	case 5:
		return "time"
	case 6:
		return "textLength"
	case 7:
		return "custom"
	default:
		return "unknown"
	}
}

// operatorTypeName maps operator type enum values to toolkit-native names.
func operatorTypeName(op int32) string {
	switch op {
	case 0:
		return "between"
	case 1:
		return "equal"
	case 2:
		return "greaterThan"
	case 3:
		return "greaterOrEqual"
	case 4:
		return "lessThan"
	case 5:
		return "lessOrEqual"
	case 6:
		return "notEqual"
	default:
		return "unknown"
	}
}

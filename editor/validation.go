package editor

import (
	"fmt"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// ValidationType names a data validation type. The constants below cover the
// common validation types; any of the engine's validation types is also
// accepted by its own name (for example "wholeNumber" or "textLength"), and
// names are matched case- and punctuation-insensitively.
type ValidationType string

// The curated validation types. Each names a category of validation rule.
const (
	ValidationTypeAnyValue    ValidationType = "anyValue"
	ValidationTypeWholeNumber ValidationType = "wholeNumber"
	ValidationTypeDecimal     ValidationType = "decimal"
	ValidationTypeList        ValidationType = "list"
	ValidationTypeDate        ValidationType = "date"
	ValidationTypeTime        ValidationType = "time"
	ValidationTypeTextLength  ValidationType = "textLength"
	ValidationTypeCustom      ValidationType = "custom"
)

// OperatorType names an operator for comparison-based validations. The
// constants below cover the common operators; any of the engine's operator
// types is also accepted by its own name (for example "between" or "greaterThan"),
// and names are matched case- and punctuation-insensitively.
type OperatorType string

// The curated operator types. Each names a comparison operator.
const (
	OperatorTypeBetween      OperatorType = "between"
	OperatorTypeEqual        OperatorType = "equal"
	OperatorTypeNotEqual     OperatorType = "notEqual"
	OperatorTypeLessThan     OperatorType = "lessThan"
	OperatorTypeLessOrEqual  OperatorType = "lessOrEqual"
	OperatorTypeGreaterThan  OperatorType = "greaterThan"
	OperatorTypeGreaterOrEqual OperatorType = "greaterOrEqual"
)

// ValidationAlertType names the alert style shown when validation fails.
type ValidationAlertType string

// The curated alert types.
const (
	ValidationAlertInformation ValidationAlertType = "information"
	ValidationAlertWarning     ValidationAlertType = "warning"
	ValidationAlertStop        ValidationAlertType = "stop"
)

// AddDataValidation creates a WorksheetAction that adds a data validation rule
// to the specified cell range. The range is an A1-style reference (e.g. "A1:A10").
// The validation is configured by the provided DataValidationActions.
//
// Parameters:
//   - range: The cell range to validate (e.g. "A1:A10", "B2:B100").
//   - actions: A variadic list of DataValidationAction functions to configure
//     the validation rule.
//
// Returns:
//   - WorksheetAction: A function that adds the validation to the worksheet.
//
// Example:
//
//	editor.InWorksheet("Sheet1",
//	    editor.AddDataValidation("A1:A10",
//	        editor.WithValidationType(editor.ValidationTypeWholeNumber),
//	        editor.WithValidationOperator(editor.OperatorTypeBetween),
//	        editor.WithValidationFormula1("1"),
//	        editor.WithValidationFormula2("100"),
//	        editor.WithValidationErrorMessage("Please enter a number between 1 and 100"),
//	    ),
//	)
func AddDataValidation(cellRange string, actions ...DataValidationAction) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		validation, err := cells.AddValidation(worksheet, cellRange)
		if err != nil {
			return err
		}
		for _, action := range actions {
			if err := action(validation); err != nil {
				return err
			}
		}
		return nil
	}
}

// InValidation creates a WorksheetAction that applies data validation actions
// to the worksheet's existing validation at the given index. Validations are
// indexed from zero in the order they were added.
//
// Parameters:
//   - index: The zero-based index of the validation to modify.
//   - actions: A variadic list of DataValidationAction functions to apply.
//
// Returns:
//   - WorksheetAction: A function that modifies the validation.
//
// Example:
//
//	editor.InWorksheet("Sheet1",
//	    editor.InValidation(0,
//	        editor.WithValidationErrorMessage("Updated message"),
//	    ),
//	)
func InValidation(index int, actions ...DataValidationAction) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		validations, err := worksheet.GetValidations()
		if err != nil {
			return err
		}
		validation, err := cells.GetValidation(validations, index)
		if err != nil {
			return err
		}
		for _, action := range actions {
			if err := action(validation); err != nil {
				return err
			}
		}
		return nil
	}
}

// DeleteValidation creates a WorksheetAction that removes the data validation
// at the given index from the worksheet.
//
// Parameters:
//   - index: The zero-based index of the validation to delete.
//
// Returns:
//   - WorksheetAction: A function that deletes the validation.
func DeleteValidation(index int) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		validations, err := worksheet.GetValidations()
		if err != nil {
			return err
		}
		count, err := validations.GetCount()
		if err != nil {
			return err
		}
		if index < 0 || int32(index) >= count {
			return fmt.Errorf("validation index %d out of range [0, %d): %w",
				index, count, toolkiterrors.ErrValidationNotFound)
		}
		// Get the validation and remove each of its areas
		v, err := validations.Get(int32(index))
		if err != nil {
			return err
		}
		areas, err := v.GetAreas()
		if err != nil {
			return err
		}
		for i := range areas {
			if err := validations.RemoveArea(&areas[i]); err != nil {
				return err
			}
		}
		return nil
	}
}

// WithValidationType creates a DataValidationAction that sets the validation type.
//
// Parameters:
//   - validationType: The type of validation (e.g. ValidationTypeWholeNumber,
//     ValidationTypeList, ValidationTypeCustom).
//
// Returns:
//   - DataValidationAction: A function that sets the validation type.
//
// Example:
//
//	editor.WithValidationType(editor.ValidationTypeList)
func WithValidationType(validationType ValidationType) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		vt, err := cells.ResolveValidationType(string(validationType))
		if err != nil {
			return err
		}
		return validation.SetType(vt)
	}
}

// WithValidationOperator creates a DataValidationAction that sets the comparison
// operator for the validation.
//
// Parameters:
//   - operator: The comparison operator (e.g. OperatorTypeBetween,
//     OperatorTypeGreaterThan).
//
// Returns:
//   - DataValidationAction: A function that sets the validation operator.
//
// Example:
//
//	editor.WithValidationOperator(editor.OperatorTypeBetween)
func WithValidationOperator(operator OperatorType) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		op, err := cells.ResolveOperatorType(string(operator))
		if err != nil {
			return err
		}
		return validation.SetOperator(op)
	}
}

// WithValidationFormula1 creates a DataValidationAction that sets the first
// formula or value for the validation. For comparison operators like "between",
// this is the lower bound. For "list" validations, this is the source range or
// comma-separated list.
//
// Parameters:
//   - formula: The formula or value string.
//
// Returns:
//   - DataValidationAction: A function that sets the first formula.
//
// Example:
//
//	editor.WithValidationFormula1("1")       // For numeric validations
//	editor.WithValidationFormula1("A1:A10")  // For list validations
func WithValidationFormula1(formula string) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		return validation.SetFormula1_String(formula)
	}
}

// WithValidationFormula2 creates a DataValidationAction that sets the second
// formula or value for the validation. This is used with operators like "between"
// to specify the upper bound.
//
// Parameters:
//   - formula: The formula or value string.
//
// Returns:
//   - DataValidationAction: A function that sets the second formula.
//
// Example:
//
//	editor.WithValidationFormula2("100")  // For "between 1 and 100"
func WithValidationFormula2(formula string) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		return validation.SetFormula2_String(formula)
	}
}

// WithValidationList creates a DataValidationAction that configures the
// validation as a dropdown list with the specified items. This is a convenience
// function that sets the validation type to List and the formula to the
// comma-separated items.
//
// Parameters:
//   - items: The list items (e.g. []string{"Yes", "No", "Maybe"}).
//
// Returns:
//   - DataValidationAction: A function that sets up a dropdown list.
//
// Example:
//
//	editor.WithValidationList([]string{"Yes", "No", "Maybe"})
func WithValidationList(items []string) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		if err := validation.SetType(asposecells.ValidationType_List); err != nil {
			return err
		}
		// Join items with comma for the formula
		formula := ""
		for i, item := range items {
			if i > 0 {
				formula += ","
			}
			formula += item
		}
		return validation.SetFormula1_String(formula)
	}
}

// WithValidationInCellDropDown creates a DataValidationAction that enables or
// disables the in-cell dropdown for list validations.
//
// Parameters:
//   - show: true to show the dropdown, false to hide it.
//
// Returns:
//   - DataValidationAction: A function that sets the dropdown visibility.
//
// Example:
//
//	editor.WithValidationInCellDropDown(true)
func WithValidationInCellDropDown(show bool) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		return validation.SetInCellDropDown(show)
	}
}

// WithValidationIgnoreBlank creates a DataValidationAction that enables or
// disables ignoring blank cells in the validation.
//
// Parameters:
//   - ignore: true to ignore blanks, false to treat them as invalid.
//
// Returns:
//   - DataValidationAction: A function that sets the blank handling.
//
// Example:
//
//	editor.WithValidationIgnoreBlank(true)
func WithValidationIgnoreBlank(ignore bool) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		return validation.SetIgnoreBlank(ignore)
	}
}

// WithValidationShowInput creates a DataValidationAction that enables or
// disables the input message when a cell is selected.
//
// Parameters:
//   - show: true to show the input message, false to hide it.
//
// Returns:
//   - DataValidationAction: A function that sets the input message visibility.
//
// Example:
//
//	editor.WithValidationShowInput(true)
func WithValidationShowInput(show bool) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		return validation.SetShowInput(show)
	}
}

// WithValidationShowError creates a DataValidationAction that enables or
// disables the error alert when invalid data is entered.
//
// Parameters:
//   - show: true to show the error alert, false to hide it.
//
// Returns:
//   - DataValidationAction: A function that sets the error alert visibility.
//
// Example:
//
//	editor.WithValidationShowError(true)
func WithValidationShowError(show bool) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		return validation.SetShowError(show)
	}
}

// WithValidationAlertStyle creates a DataValidationAction that sets the alert
// style shown when validation fails.
//
// Parameters:
//   - alertType: The alert style (e.g. ValidationAlertStop, ValidationAlertWarning,
//     ValidationAlertInformation).
//
// Returns:
//   - DataValidationAction: A function that sets the alert style.
//
// Example:
//
//	editor.WithValidationAlertStyle(editor.ValidationAlertStop)
func WithValidationAlertStyle(alertType ValidationAlertType) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		at, err := cells.ResolveValidationAlertType(string(alertType))
		if err != nil {
			return err
		}
		return validation.SetAlertStyle(at)
	}
}

// WithValidationErrorTitle creates a DataValidationAction that sets the title
// of the error alert dialog.
//
// Parameters:
//   - title: The error dialog title.
//
// Returns:
//   - DataValidationAction: A function that sets the error title.
//
// Example:
//
//	editor.WithValidationErrorTitle("Invalid Input")
func WithValidationErrorTitle(title string) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		return validation.SetErrorTitle(title)
	}
}

// WithValidationErrorMessage creates a DataValidationAction that sets the
// message of the error alert dialog.
//
// Parameters:
//   - message: The error message.
//
// Returns:
//   - DataValidationAction: A function that sets the error message.
//
// Example:
//
//	editor.WithValidationErrorMessage("Please enter a valid number between 1 and 100")
func WithValidationErrorMessage(message string) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		return validation.SetErrorMessage(message)
	}
}

// WithValidationInputTitle creates a DataValidationAction that sets the title
// of the input message.
//
// Parameters:
//   - title: The input message title.
//
// Returns:
//   - DataValidationAction: A function that sets the input title.
//
// Example:
//
//	editor.WithValidationInputTitle("Enter Value")
func WithValidationInputTitle(title string) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		return validation.SetInputTitle(title)
	}
}

// WithValidationInputMessage creates a DataValidationAction that sets the
// input message shown when a cell is selected.
//
// Parameters:
//   - message: The input message.
//
// Returns:
//   - DataValidationAction: A function that sets the input message.
//
// Example:
//
//	editor.WithValidationInputMessage("Please enter a number between 1 and 100")
func WithValidationInputMessage(message string) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		return validation.SetInputMessage(message)
	}
}

// WithDataValidationRange creates a DataValidationAction that configures the
// validation as a numeric range check. This is a convenience function that sets
// the validation type to WholeNumber or Decimal, the operator to Between, and
// the formulas to the min and max values.
//
// Parameters:
//   - min: The minimum allowed value (inclusive).
//   - max: The maximum allowed value (inclusive).
//   - allowDecimal: If true, allows decimal values; if false, only integers.
//
// Returns:
//   - DataValidationAction: A function that sets up a numeric range validation.
//
// Example:
//
//	editor.WithDataValidationRange(1, 100, false)  // Integer 1-100
//	editor.WithDataValidationRange(0.0, 1.0, true) // Decimal 0.0-1.0
func WithDataValidationRange(min, max float64, allowDecimal bool) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		// Set validation type
		if allowDecimal {
			if err := validation.SetType(asposecells.ValidationType_Decimal); err != nil {
				return err
			}
		} else {
			if err := validation.SetType(asposecells.ValidationType_WholeNumber); err != nil {
				return err
			}
		}
		// Set operator to Between
		if err := validation.SetOperator(asposecells.OperatorType_Between); err != nil {
			return err
		}
		// Set formulas
		minStr := fmt.Sprintf("%v", min)
		maxStr := fmt.Sprintf("%v", max)
		if err := validation.SetFormula1_String(minStr); err != nil {
			return err
		}
		return validation.SetFormula2_String(maxStr)
	}
}

// WithDataValidationLength creates a DataValidationAction that configures the
// validation as a text length check. This is a convenience function that sets
// the validation type to TextLength and the operator to Between with the min
// and max length values.
//
// Parameters:
//   - min: The minimum allowed length (inclusive).
//   - max: The maximum allowed length (inclusive).
//
// Returns:
//   - DataValidationAction: A function that sets up a text length validation.
//
// Example:
//
//	editor.WithDataValidationLength(3, 50)  // Text length 3-50 characters
func WithDataValidationLength(min, max int) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		// Set validation type
		if err := validation.SetType(asposecells.ValidationType_TextLength); err != nil {
			return err
		}
		// Set operator to Between
		if err := validation.SetOperator(asposecells.OperatorType_Between); err != nil {
			return err
		}
		// Set formulas
		minStr := fmt.Sprintf("%d", min)
		maxStr := fmt.Sprintf("%d", max)
		if err := validation.SetFormula1_String(minStr); err != nil {
			return err
		}
		return validation.SetFormula2_String(maxStr)
	}
}

// WithDataValidationFormula creates a DataValidationAction that configures the
// validation as a custom formula check. This is a convenience function that sets
// the validation type to Custom and the formula to the provided expression.
// The formula should return TRUE for valid values and FALSE for invalid values.
//
// Parameters:
//   - formula: The custom validation formula (e.g., "=A1>0", "=AND(A1>0,A1<100)").
//
// Returns:
//   - DataValidationAction: A function that sets up a custom formula validation.
//
// Example:
//
//	editor.WithDataValidationFormula("=A1>0")                    // Positive numbers only
//	editor.WithDataValidationFormula("=AND(A1>0,A1<100)")       // Between 0 and 100
//	editor.WithDataValidationFormula("=ISNUMBER(SEARCH(\"@\",A1))") // Email-like
func WithDataValidationFormula(formula string) DataValidationAction {
	return func(validation *asposecells.Validation) error {
		// Set validation type to Custom
		if err := validation.SetType(asposecells.ValidationType_Custom); err != nil {
			return err
		}
		// Set the formula
		return validation.SetFormula1_String(formula)
	}
}

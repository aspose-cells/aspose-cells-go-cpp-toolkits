// Package editor provides functionality for editing spreadsheet documents
// using a fluent, action-based Domain Specific Language (DSL).
package editor

import (
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// WithActiveSheet creates a WorkbookAction that activates the worksheet
// with the given name.
//
// Parameters:
//   - sheetName: The name of the worksheet to activate.
//
// Returns:
//   - WorkbookAction: A function that modifies the workbook by setting
//     the active sheet.
func WithActiveSheet(sheetName string) WorkbookAction {
	return func(workbook *asposecells.Workbook) error {
		wss, err := workbook.GetWorksheets()
		if err != nil {
			return err
		}
		return wss.SetActiveSheetName(sheetName)
	}
}

// WithAddWorksheet creates a WorkbookAction that add the worksheet
//
// Parameters:
//   - sheetName: The name of the worksheet to add.
//
// Returns:
//   - WorkbookAction: A function that adds the workbook by the sheet name.

func WithAddWorksheet(newSheetName string) WorkbookAction {
	return func(workbook *asposecells.Workbook) error {
		wss, err := workbook.GetWorksheets()
		if err != nil {
			return err
		}
		_, err = wss.Add_String(newSheetName)
		return err
	}
}

// WithDeleteWorksheet creates a WorkbookAction that delete the worksheet
//
// Parameters:
//   - sheetName: The name of the worksheet to delete.
//
// Returns:
//   - WorkbookAction: A function that deletes the workbook by the sheet name.
func WithDeleteWorksheet(sheetName string) WorkbookAction {
	return func(workbook *asposecells.Workbook) error {
		wss, err := workbook.GetWorksheets()
		if err != nil {
			return err
		}
		return wss.RemoveAt_String(sheetName)
	}
}

// WithRenameWorksheet creates a WorkbookAction that renames an existing worksheet.
//
// Parameters:
//   - sheetName: The current name of the worksheet to rename.
//   - newSheetName: The new name to assign to the worksheet.
//
// Returns:
//   - WorkbookAction: A function that renames the specified worksheet.
func WithRenameWorksheet(sheetName string, newSheetName string) WorkbookAction {
	return func(workbook *asposecells.Workbook) error {
		wss, err := workbook.GetWorksheets()
		if err != nil {
			return err
		}
		ws, err := cells.WorksheetByName(wss, sheetName)
		if err != nil {
			return err
		}
		return ws.SetName(newSheetName)
	}
}

// InDefaultStyle creates a WorkbookAction that targets the workbook's default style.
// It retrieves the current default style and applies a sequence of StyleActions to it.
// This follows a "Fail-Fast" principle; if any action fails, execution stops.
//
// Parameters:
//   - actions: A variadic list of StyleAction functions to be applied sequentially
//     to the workbook's default style.
//
// Returns:
//   - WorkbookAction: A function that modifies the default style of the workbook.
func InDefaultStyle(actions ...StyleAction) WorkbookAction {
	return func(workbook *asposecells.Workbook) error {
		defaultStyle, err := workbook.GetDefaultStyle()
		if err != nil {
			return err
		}
		for _, action := range actions {
			if err = action(defaultStyle); err != nil {
				return err
			}
		}
		return nil
	}
}

// CalculateAll creates a WorkbookAction that recalculates every formula in the
// workbook. Place it after the actions that set or depend on formula values so
// the saved workbook holds current results.
//
// Returns:
//   - WorkbookAction: A function that recalculates the workbook.
func CalculateAll() WorkbookAction {
	return func(workbook *asposecells.Workbook) error {
		return workbook.CalculateFormula()
	}
}

// InWorksheet creates a WorkbookAction that scopes subsequent operations to a specific worksheet.
// It resolves the target worksheet using a flexible identifier (e.g., sheet name or index),
// and then executes a series of WorksheetActions within that context.
//
// Parameters:
//
//   - identifier: An interface{} value used to locate the target worksheet.
//     Supported types typically include string (sheet name) or int (sheet index).
//
//   - actions: A variadic list of WorksheetAction functions to be executed
//     on the resolved worksheet.
//
// Returns:
//   - WorkbookAction: A function that targets a specific worksheet and executes
//     nested actions. If the worksheet cannot be resolved, an error is returned.
func InWorksheet(identifier interface{}, actions ...WorksheetAction) WorkbookAction {
	return func(wb *asposecells.Workbook) error {
		sheet, err := resolveWorksheet(wb, identifier)
		if err != nil {
			return err
		}

		for _, action := range actions {
			if err = action(sheet); err != nil {
				return err
			}
		}
		return nil
	}
}

// Package editor provides functionality for editing spreadsheet documents
// using a fluent, action-based Domain Specific Language (DSL).
package editor

import asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"

//func WithDefaultStyle(styleAction StyleAction) WorkbookAction {
//	return func(workbook *asposecells.Workbook) error {
//		workbook.SetDefaultStyle(defaultStyle)
//		return nil
//	}
//}

// WithActiveSheet creates a WorkbookAction that activates the worksheet
// at the specified zero-based index.
//
// Parameters:
//   - sheetIndex: The zero-based index of the worksheet to activate.
//
// Returns:
//   - WorkbookAction: A function that modifies the workbook by setting
//     the active sheet index.
func WithActiveSheet(sheetIndex int) WorkbookAction {
	return func(workbook *asposecells.Workbook) error {
		wss, _ := workbook.GetWorksheets()
		wss.SetActiveSheetIndex(int32(sheetIndex))
		return nil
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

//func InChart(sheetID interface{}, chartIndex int, actions ...ChartAction) WorkbookAction {
//	return func(wb *asposecells.Workbook) error {
//		sheet, _ := resolveSheet(wb, sheetID)
//		charts, _ := sheet.GetCharts()
//		chart, _ := charts.Get_Int(int32(chartIndex))
//		for _, action := range actions {
//			if err := action(chart); err != nil {
//				return err
//			}
//		}
//		return nil
//	}
//}

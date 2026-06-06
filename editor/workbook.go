package editor

import asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"

//func WithDefaultFont(fontName string, size int32) WorkbookAction {
//	return func(workbook *asposecells.Workbook) error {
//		defaultStyle, _ := workbook.GetDefaultStyle()
//		font, _ := defaultStyle.GetFont()
//		font.SetName_String(fontName)
//		font.SetSize(size)
//		return nil
//	}
//}

func WithActiveSheet(sheetIndex int) WorkbookAction {
	return func(workbook *asposecells.Workbook) error {
		wss, _ := workbook.GetWorksheets()
		wss.SetActiveSheetIndex(int32(sheetIndex))
		return nil
	}
}

func InSheet(identifier interface{}, actions ...WorksheetAction) WorkbookAction {
	return func(wb *asposecells.Workbook) error {
		sheet, err := resolveSheet(wb, identifier)
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

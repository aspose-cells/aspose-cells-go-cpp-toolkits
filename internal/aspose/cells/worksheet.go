package cells

import (
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

func GetCellAreaWithWorksheet(workbook *asposecells.Workbook, worksheet string) (int32, *asposecells.CellArea, error) {
	worksheets, err := workbook.GetWorksheets()
	if err == nil {
		worksheet, err := worksheets.Get_String(worksheet)
		sheetIndex, _ := worksheet.GetIndex()
		if err == nil {
			cells, _ := worksheet.GetCells()
			dataColumnIndex, _ := cells.GetMaxDataColumn()
			dataRowIndex, _ := cells.GetMaxDataRow()
			endCellName, _ := asposecells.CellsHelper_CellIndexToName(dataRowIndex, dataColumnIndex)
			cellArea, _ := asposecells.CellArea_CreateCellArea_String_String("A1", endCellName)
			return sheetIndex, cellArea, nil
		}
		return sheetIndex, nil, err
	}
	return -1, nil, err
}

func GetCellsWithWorksheet(workbook *asposecells.Workbook, worksheet string) (*asposecells.Cells, error) {
	worksheets, err := workbook.GetWorksheets()
	if err != nil {
		return nil, err
	}
	ws, err := worksheets.Get_String(worksheet)
	if err != nil {
		return nil, err
	}
	cells, err := ws.GetCells()
	if err != nil {
		return nil, err
	}
	return cells, err
}

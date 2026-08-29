package cells

import (
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

func GetCellAreaWithWorksheet(workbook *asposecells.Workbook, worksheet string) (int32, *asposecells.CellArea, error) {
	worksheets, err := workbook.GetWorksheets()
	if err != nil {
		return -1, nil, err
	}
	ws, err := worksheets.Get_String(worksheet)
	if err != nil {
		return -1, nil, err
	}
	sheetIndex, err := ws.GetIndex()
	if err != nil {
		return -1, nil, err
	}
	cells, err := ws.GetCells()
	if err != nil {
		return -1, nil, err
	}
	dataColumnIndex, err := cells.GetMaxDataColumn()
	if err != nil {
		return -1, nil, err
	}
	dataRowIndex, err := cells.GetMaxDataRow()
	if err != nil {
		return -1, nil, err
	}
	endCellName, err := asposecells.CellsHelper_CellIndexToName(dataRowIndex, dataColumnIndex)
	if err != nil {
		return -1, nil, err
	}
	cellArea, err := asposecells.CellArea_CreateCellArea_String_String("A1", endCellName)
	if err != nil {
		return -1, nil, err
	}
	return sheetIndex, cellArea, nil
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

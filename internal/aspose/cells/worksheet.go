package cells

import (
	"fmt"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// WorksheetByName returns the worksheet in the collection whose name equals name.
//
// It resolves by iterating the collection (Get_Int + GetName) instead of calling
// Get_String, because the binding's Get_String returns err=nil together with a
// DANGLING native handle when the name does not exist; any subsequent use of that
// handle crashes the whole process. Iterating makes a missing name a normal error.
func WorksheetByName(worksheets *asposecells.WorksheetCollection, name string) (*asposecells.Worksheet, error) {
	count, err := worksheets.GetCount()
	if err != nil {
		return nil, err
	}
	for i := int32(0); i < count; i++ {
		ws, err := worksheets.Get_Int(i)
		if err != nil {
			return nil, err
		}
		got, err := ws.GetName()
		if err != nil {
			return nil, err
		}
		if got == name {
			return ws, nil
		}
	}
	return nil, fmt.Errorf("worksheet %q not found: %w", name, toolkiterrors.ErrWorksheetNotFound)
}

func GetCellAreaWithWorksheet(workbook *asposecells.Workbook, worksheet string) (int32, *asposecells.CellArea, error) {
	worksheets, err := workbook.GetWorksheets()
	if err != nil {
		return -1, nil, err
	}
	ws, err := WorksheetByName(worksheets, worksheet)
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
	ws, err := WorksheetByName(worksheets, worksheet)
	if err != nil {
		return nil, err
	}
	cells, err := ws.GetCells()
	if err != nil {
		return nil, err
	}
	return cells, err
}

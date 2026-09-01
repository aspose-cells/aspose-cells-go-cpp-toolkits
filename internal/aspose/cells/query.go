package cells

import (
	"fmt"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// GetCell returns the cell at the zero-based (row, col) coordinates.
func GetCell(ws *asposecells.Worksheet, row, col int32) (*asposecells.Cell, error) {
	cs, err := ws.GetCells()
	if err != nil {
		return nil, err
	}
	return cs.Get_Int_Int(row, col)
}

// UsedRange returns the dimensions of the worksheet's used range: the maximum
// used row plus one and the maximum used column plus one. An empty sheet
// yields (0, 0).
func UsedRange(ws *asposecells.Worksheet) (rows, cols int32, err error) {
	cs, err := ws.GetCells()
	if err != nil {
		return 0, 0, err
	}
	if rows, err = cs.GetMaxDataRow(); err != nil {
		return 0, 0, err
	}
	if cols, err = cs.GetMaxDataColumn(); err != nil {
		return 0, 0, err
	}
	return rows + 1, cols + 1, nil
}

// MergedAreas returns the merged cell regions of the worksheet as Areas. The
// regions come from the engine's Cells.GetMergedAreas, so each anchor region
// is reported exactly once.
func MergedAreas(ws *asposecells.Worksheet) ([]Area, error) {
	cs, err := ws.GetCells()
	if err != nil {
		return nil, err
	}
	cellAreas, err := cs.GetMergedAreas()
	if err != nil {
		return nil, err
	}
	areas := make([]Area, 0, len(cellAreas))
	for _, ca := range cellAreas {
		startRow, err := ca.Get_StartRow()
		if err != nil {
			return nil, err
		}
		endRow, err := ca.Get_EndRow()
		if err != nil {
			return nil, err
		}
		startCol, err := ca.Get_StartColumn()
		if err != nil {
			return nil, err
		}
		endCol, err := ca.Get_EndColumn()
		if err != nil {
			return nil, err
		}
		areas = append(areas, Area{
			Start: CellRef{Row: int(startRow), Col: int(startCol)},
			End:   CellRef{Row: int(endRow), Col: int(endCol)},
		})
	}
	return areas, nil
}

// SheetNames returns the names of all worksheets in the workbook, in order.
func SheetNames(wb *asposecells.Workbook) ([]string, error) {
	wss, err := wb.GetWorksheets()
	if err != nil {
		return nil, err
	}
	count, err := wss.GetCount()
	if err != nil {
		return nil, err
	}
	names := make([]string, 0, count)
	for i := int32(0); i < count; i++ {
		ws, err := wss.Get_Int(i)
		if err != nil {
			return nil, err
		}
		name, err := ws.GetName()
		if err != nil {
			return nil, err
		}
		names = append(names, name)
	}
	return names, nil
}

// SheetByIndex returns the worksheet at the given zero-based index. Out-of-range
// indexes return ErrInvalidSheetID instead of relying on the engine's
// Get_Int behavior.
func SheetByIndex(wb *asposecells.Workbook, index int) (*asposecells.Worksheet, error) {
	wss, err := wb.GetWorksheets()
	if err != nil {
		return nil, err
	}
	count, err := wss.GetCount()
	if err != nil {
		return nil, err
	}
	if index < 0 || int32(index) >= count {
		return nil, fmt.Errorf("sheet index %d out of range (workbook has %d worksheets): %w", index, count, toolkiterrors.ErrInvalidSheetID)
	}
	return wss.Get_Int(int32(index))
}

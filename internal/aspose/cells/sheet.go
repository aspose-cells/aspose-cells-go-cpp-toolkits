package cells

import (
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Sheet selects a worksheet by index or by name. Both query and transfer
// embed it in their options, so the two read/write packages expose the same
// sheet-selection vocabulary (WithSheet / WithSheetIndex).
//
// Index selection is the default everywhere: it is immune to the
// evaluation-mode load-time name corruption, because resolving by position
// never reads a name. Name selection matches an explicit worksheet name and
// can fail with ErrWorksheetNotFound if the engine corrupted that name at
// load time.
type Sheet struct {
	UseIndex bool
	Index    int
	Name     string
}

// FirstSheet selects the first worksheet by index. It is the default target
// for query and transfer.
var FirstSheet = Sheet{UseIndex: true}

// Resolve returns the selected worksheet of wb.
func (s Sheet) Resolve(wb *asposecells.Workbook) (*asposecells.Worksheet, error) {
	if s.UseIndex {
		return SheetByIndex(wb, s.Index)
	}
	wss, err := wb.GetWorksheets()
	if err != nil {
		return nil, err
	}
	return WorksheetByName(wss, s.Name)
}

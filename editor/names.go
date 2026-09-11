package editor

import (
	"fmt"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// SetCellComment creates a WorksheetAction that adds or replaces the comment
// note on a specific cell. An existing comment on the cell is overwritten.
//
// Parameters:
//   - row: The zero-based row index of the target cell.
//   - column: The zero-based column index of the target cell.
//   - text: The comment's note text.
//
// Returns:
//   - WorksheetAction: A function that writes the comment. Read it back with
//     query.ReadCellComment.
func SetCellComment(row, column int, text string) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		return cells.SetCellComment(worksheet, int32(row), int32(column), text)
	}
}

// ClearComments creates a WorksheetAction that removes every comment on the
// worksheet.
//
// Returns:
//   - WorksheetAction: A function that clears the sheet's comments.
func ClearComments() WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		return cells.ClearComments(worksheet)
	}
}

// DefineNamedRange creates a WorksheetAction that defines a workbook-level
// named range referring to a block of cells on the applied worksheet. If a name
// with the same text already exists, its reference is updated instead, so the
// action is idempotent.
//
// The reference is stored as "='SheetName'!$A$1:$B$2". List the workbook's
// names with query.NamedRanges and read one back with query.ReadNamedRange.
//
// Parameters:
//   - name: The name to define, e.g. "MyRange".
//   - startRow: The zero-based row of the range's first cell.
//   - startColumn: The zero-based column of the range's first cell.
//   - endRow: The zero-based row of the range's last cell.
//   - endColumn: The zero-based column of the range's last cell.
//
// Returns:
//   - WorksheetAction: A function that defines the name. Returns ErrInvalidRange
//     when the start cell lies below or to the right of the end cell.
func DefineNamedRange(name string, startRow, startColumn, endRow, endColumn int) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		if startRow > endRow || startColumn > endColumn {
			return fmt.Errorf("invalid range for named %q: %w", name, toolkiterrors.ErrInvalidRange)
		}
		sheetName, err := worksheet.GetName()
		if err != nil {
			return err
		}
		refersTo := "='" + sheetName + "'!" +
			(cells.CellRef{Row: startRow, Col: startColumn}).AbsoluteString() + ":" +
			(cells.CellRef{Row: endRow, Col: endColumn}).AbsoluteString()
		wb, err := worksheet.GetWorkbook()
		if err != nil {
			return err
		}
		return cells.SetOrAddNamedRange(wb, name, refersTo)
	}
}

// Encrypt creates a WorkbookAction that encrypts the workbook so the saved
// file requires password to open. It sets the workbook's encryption password
// and selects strong (AES) encryption.
//
// Decrypting an encrypted workbook requires supplying the password at load
// time, which the toolkit's loader does not yet expose; to read an encrypted
// file back, load it with a password through the underlying engine.
//
// Parameters:
//   - password: The password the saved file will require to open.
//
// Returns:
//   - WorkbookAction: A function that encrypts the workbook before it is saved.
func Encrypt(password string) WorkbookAction {
	return func(workbook *asposecells.Workbook) error {
		return cells.EncryptWorkbook(workbook, password)
	}
}

package cells

import (
	"strings"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// WorkbookNames returns the workbook's collection of defined names. Named
// ranges are workbook-scoped (the names live on the WorksheetCollection), even
// though each refers to cells on a specific worksheet.
func WorkbookNames(wb *asposecells.Workbook) (*asposecells.NameCollection, error) {
	wss, err := wb.GetWorksheets()
	if err != nil {
		return nil, err
	}
	return wss.GetNames()
}

// FindName returns the workbook name whose text equals name (case-insensitive,
// matching Excel semantics), or wraps ErrNameNotFound. It iterates the
// collection instead of calling NameCollection.Get_String because, like the
// worksheet lookup, the engine returns a dangling handle for a missing name
// rather than an error.
func FindName(wb *asposecells.Workbook, name string) (*asposecells.Name, error) {
	names, err := WorkbookNames(wb)
	if err != nil {
		return nil, err
	}
	count, err := names.GetCount()
	if err != nil {
		return nil, err
	}
	for i := int32(0); i < count; i++ {
		n, err := names.Get_Int(i)
		if err != nil {
			return nil, err
		}
		text, err := n.GetText()
		if err != nil {
			return nil, err
		}
		if strings.EqualFold(text, name) {
			return n, nil
		}
	}
	return nil, toolkiterrors.ErrNameNotFound
}

// SetOrAddNamedRange creates the named range name referring to refersTo, or
// updates the refersTo of an existing name with the same text. Overwriting an
// existing name makes the action idempotent: applying it twice leaves a single
// name.
func SetOrAddNamedRange(wb *asposecells.Workbook, name, refersTo string) error {
	names, err := WorkbookNames(wb)
	if err != nil {
		return err
	}
	count, err := names.GetCount()
	if err != nil {
		return err
	}
	for i := int32(0); i < count; i++ {
		n, err := names.Get_Int(i)
		if err != nil {
			return err
		}
		text, err := n.GetText()
		if err != nil {
			return err
		}
		if strings.EqualFold(text, name) {
			return n.SetRefersTo_String(refersTo)
		}
	}
	idx, err := names.Add(name)
	if err != nil {
		return err
	}
	n, err := names.Get_Int(idx)
	if err != nil {
		return err
	}
	return n.SetRefersTo_String(refersTo)
}

// SetCellComment adds or replaces the comment note on the cell at (row, col).
func SetCellComment(ws *asposecells.Worksheet, row, col int32, text string) error {
	comments, err := ws.GetComments()
	if err != nil {
		return err
	}
	idx, err := comments.Add_Int_Int(row, col)
	if err != nil {
		return err
	}
	comment, err := comments.Get_Int(idx)
	if err != nil {
		return err
	}
	return comment.SetNote(text)
}

// CellComment returns the comment note on the cell at (row, col), or the empty
// string when the cell has no comment. It iterates the collection and matches
// on the comment's position, because the binding's Cell.GetComment returns a
// dangling handle (a hard crash on use) for a comment-free cell.
func CellComment(ws *asposecells.Worksheet, row, col int32) (string, error) {
	comments, err := ws.GetComments()
	if err != nil {
		return "", err
	}
	count, err := comments.GetCount()
	if err != nil {
		return "", err
	}
	for i := int32(0); i < count; i++ {
		comment, err := comments.Get_Int(i)
		if err != nil {
			return "", err
		}
		crow, err := comment.GetRow()
		if err != nil {
			return "", err
		}
		ccol, err := comment.GetColumn()
		if err != nil {
			return "", err
		}
		if crow == row && ccol == col {
			return comment.GetNote()
		}
	}
	return "", nil
}

// ClearComments removes every comment on the worksheet.
func ClearComments(ws *asposecells.Worksheet) error {
	return ws.ClearComments()
}

// EncryptWorkbook sets the workbook's encryption password so the saved file
// requires it to open, and selects strong (AES) encryption. A later save via
// WorkbookToByteData produces an encrypted output. To read the file back, a
// password must be supplied at load time (LoadOptions.Password), which the
// toolkit's shared loader does not yet expose.
func EncryptWorkbook(wb *asposecells.Workbook, password string) error {
	settings, err := wb.GetSettings()
	if err != nil {
		return err
	}
	if err := settings.SetPassword(password); err != nil {
		return err
	}
	return wb.SetEncryptionOptions(asposecells.EncryptionType_StrongCryptographicProvider, 128)
}

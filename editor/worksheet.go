package editor

import (
	"fmt"
	"strings"
	"time"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// SetCellValue creates a WorksheetAction that sets the value of a specific cell.
// It uses a type switch to handle various data types and converts them into
// Aspose.Cells compatible objects before writing.
//
// Parameters:
//   - row: The zero-based row index of the target cell.
//   - column: The zero-based column index of the target cell.
//   - value: The value to set. Supported types include int8, uint16, uint64, int16,
//     int32, int, int64, float32, float64, string, bool, and time.Time.
//
// Returns:
//   - WorksheetAction: A function that modifies the specified cell. Returns an error
//     if the provided value type is unsupported.
func SetCellValue(row, column int, value interface{}) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, err := worksheet.GetCells()
		if err != nil {
			return err
		}
		cell, err := cells.Get_Int_Int(int32(row), int32(column))
		if err != nil {
			return err
		}
		obj, err := toObject(value)
		if err != nil {
			return err
		}
		return cell.PutValue_Object(obj)
	}
}

// SetValue creates a WorksheetAction that sets the value for all cells within a specified range.
// Similar to SetCellValue, it handles multiple data types via a type switch.
//
// Parameters:
//   - beginRow: The zero-based starting row index.
//   - beginColumn: The zero-based starting column index.
//   - rows: The number of rows in the range.
//   - columns: The number of columns in the range.
//   - value: The value to apply across the range. Supports the same types as SetCellValue.
//
// Returns:
//   - WorksheetAction: A function that populates the range with the given value.
//     Returns an error if the value type is unsupported.
func SetValue(beginRow, beginColumn, rows, columns int, value interface{}) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, err := worksheet.GetCells()
		if err != nil {
			return err
		}
		cellsRange, err := cells.CreateRange_Int_Int_Int_Int(int32(beginRow), int32(beginColumn), int32(rows), int32(columns))
		if err != nil {
			return err
		}
		obj, err := toObject(value)
		if err != nil {
			return err
		}
		return cellsRange.SetValue(obj)
	}
}

// toObject converts a supported Go value into the engine's Object type so it
// can be written to a cell or range. Supported types are the integer and
// floating-point sizes, string, bool, and time.Time.
func toObject(value interface{}) (*asposecells.Object, error) {
	switch v := value.(type) {
	case int8:
		return asposecells.NewObject_Integer8(v)
	case uint16:
		return asposecells.NewObject_UInteger16(v)
	case uint64:
		return asposecells.NewObject_ULong(v)
	case int16:
		return asposecells.NewObject_Int16(v)
	case int32:
		return asposecells.NewObject_Int(v)
	case int:
		return asposecells.NewObject_Int64(int64(v))
	case int64:
		return asposecells.NewObject_Int64(v)
	case float32:
		return asposecells.NewObject_Float(v)
	case float64:
		return asposecells.NewObject_Double(v)
	case string:
		return asposecells.NewObject_String(v)
	case bool:
		return asposecells.NewObject_Bool(v)
	case time.Time:
		return asposecells.NewObject_Date(v)
	default:
		return nil, fmt.Errorf("invalid value %v: %w", value, toolkiterrors.ErrInvalidValue)
	}
}

// SetStyle creates a WorksheetAction that applies formatting to a specific range of cells.
// It retrieves the current style, applies a sequence of StyleActions, and then assigns
// the modified style back to the target range. Follows a "Fail-Fast" principle.
//
// Parameters:
//   - beginRow: The zero-based starting row index.
//   - beginColumn: The zero-based starting column index.
//   - rows: The number of rows in the range.
//   - columns: The number of columns in the range.
//   - actions: A variadic list of StyleAction functions to modify the style properties.
//
// Returns:
//   - WorksheetAction: A function that updates the formatting of the specified range.
func SetStyle(beginRow, beginColumn, rows, columns int, actions ...StyleAction) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, err := worksheet.GetCells()
		cellsStyle, err := cells.GetStyle()
		if err != nil {
			return err
		}
		for _, action := range actions {
			if err = action(cellsStyle); err != nil {
				return err
			}
		}
		_range, err := cells.CreateRange_Int_Int_Int_Int(int32(beginRow), int32(beginColumn), int32(rows), int32(columns))
		err = _range.SetStyle_Style(cellsStyle)
		if err != nil {
			return err
		}
		return nil
	}
}

// Merge creates a WorksheetAction that merges a rectangular region of cells.
//
// Parameters:
//   - beginRow: The zero-based starting row index.
//   - beginColumn: The zero-based starting column index.
//   - rows: The number of rows to merge.
//   - columns: The number of columns to merge.
//   - mergeConflict: If true, allows merging even if it causes conflicts with existing merges.
//
// Returns:
//   - WorksheetAction: A function that performs the cell merge operation.
func Merge(beginRow, beginColumn, rows, columns int, mergeConflict bool) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, err := worksheet.GetCells()
		if err != nil {
			return err
		}
		return cells.Merge_Int_Int_Int_Int_Bool(int32(beginRow), int32(beginColumn), int32(rows), int32(columns), mergeConflict)
	}
}

// UnMerge creates a WorksheetAction that unmerges a previously merged rectangular region of cells.
//
// Parameters:
//   - beginRow: The zero-based starting row index.
//   - beginColumn: The zero-based starting column index.
//   - rows: The number of rows in the merged region.
//   - columns: The number of columns in the merged region.
//
// Returns:
//   - WorksheetAction: A function that splits the merged cells back into individual cells.
func UnMerge(beginRow, beginColumn, rows, columns int) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, err := worksheet.GetCells()
		if err != nil {
			return err
		}
		return cells.UnMerge(int32(beginRow), int32(beginColumn), int32(rows), int32(columns))
	}
}

// InsertRows creates a WorksheetAction that inserts new rows into the worksheet.
//
// Parameters:
//   - beginRow: The zero-based index at which to start inserting rows.
//   - rows: The number of rows to insert.
//   - updateReference: If true, updates references (e.g., formulas, named ranges) in other sheets.
//
// Returns:
//   - WorksheetAction: A function that adds the specified rows.
func InsertRows(beginRow int, rows int, updateReference bool) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, err := worksheet.GetCells()
		if err != nil {
			return err
		}
		return cells.InsertRows_Int_Int_Bool(int32(beginRow), int32(rows), updateReference)
	}
}

// InsertColumns creates a WorksheetAction that inserts new columns into the worksheet.
//
// Parameters:
//   - beginColumn: The zero-based index at which to start inserting columns.
//   - columns: The number of columns to insert.
//   - updateReference: If true, updates references in other parts of the workbook.
//
// Returns:
//   - WorksheetAction: A function that adds the specified columns.
func InsertColumns(beginColumn int, columns int, updateReference bool) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, err := worksheet.GetCells()
		if err != nil {
			return err
		}
		return cells.InsertColumns_Int_Int_Bool(int32(beginColumn), int32(columns), updateReference)
	}
}

// ClearContents creates a WorksheetAction that clears the values/formulas within a specified range,
// while preserving the existing cell formatting.
//
// Parameters:
//   - beginRow: The zero-based starting row index.
//   - beginColumn: The zero-based starting column index.
//   - rows: The number of rows in the range.
//   - columns: The number of columns in the range.
//
// Returns:
//   - WorksheetAction: A function that empties the contents of the target range.
func ClearContents(beginRow, beginColumn, rows, columns int) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, err := worksheet.GetCells()
		if err != nil {
			return err
		}
		cellsRange, err := cells.CreateRange_Int_Int_Int_Int(int32(beginRow), int32(beginColumn), int32(rows), int32(columns))
		if err != nil {
			return err
		}
		return cellsRange.ClearContents()
	}
}

// ClearFormats creates a WorksheetAction that clears the formatting within a specified range,
// while preserving the cell values.
//
// Parameters:
//   - beginRow: The zero-based starting row index.
//   - beginColumn: The zero-based starting column index.
//   - rows: The number of rows in the range.
//   - columns: The number of columns in the range.
//
// Returns:
//   - WorksheetAction: A function that resets the styles of the target range.
func ClearFormats(beginRow, beginColumn, rows, columns int) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, err := worksheet.GetCells()
		if err != nil {
			return err
		}
		cellsRange, err := cells.CreateRange_Int_Int_Int_Int(int32(beginRow), int32(beginColumn), int32(rows), int32(columns))
		if err != nil {
			return err
		}
		return cellsRange.ClearFormats()
	}
}

// DeleteRows creates a WorksheetAction that removes rows from the worksheet.
//
// Parameters:
//   - beginRow: The zero-based index of the first row to delete.
//   - rows: The number of rows to remove.
//   - updateReference: If true, updates references in other parts of the workbook.
//
// Returns:
//   - WorksheetAction: A function that deletes the specified rows.
func DeleteRows(beginRow int, rows int, updateReference bool) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, err := worksheet.GetCells()
		if err != nil {
			return err
		}
		_, err = cells.DeleteRows_Int_Int_Bool(int32(beginRow), int32(rows), updateReference)
		return err
	}
}

// DeleteBlankRows creates a WorksheetAction that removes all completely blank rows
// from the worksheet. A row is considered blank if it contains no data or formatting.
//
// Returns:
//   - WorksheetAction: A function that modifies the worksheet by deleting empty rows.
func DeleteBlankRows() WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, err := worksheet.GetCells()
		if err != nil {
			return err
		}
		return cells.DeleteBlankRows()
	}
}

// DeleteRange creates a WorksheetAction that deletes a specific range of cells
// and shifts the surrounding cells to fill the empty space.
//
// Parameters:
//   - beginRow: The zero-based starting row index of the range to delete.
//   - beginColumn: The zero-based starting column index of the range to delete.
//   - rows: The number of rows in the range to delete.
//   - columns: The number of columns in the range to delete.
//   - shiftType: A string specifying the direction in which remaining cells should shift.
//     Supported values are "up", "down", "left", and "right" (case-insensitive).
//     If an unrecognized value is provided, it defaults to no shift (ShiftType_None).
//
// Returns:
//   - WorksheetAction: A function that deletes the specified cell range and applies the shift.
func DeleteRange(beginRow, beginColumn, rows, columns int, shiftType string) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, err := worksheet.GetCells()
		if err != nil {
			return err
		}
		shift := asposecells.ShiftType_None
		shiftType = strings.ToLower(shiftType)
		if shiftType == "up" {
			shift = asposecells.ShiftType_Up
		}
		if shiftType == "down" {
			shift = asposecells.ShiftType_Down
		}
		if shiftType == "left" {
			shift = asposecells.ShiftType_Left
		}
		if shiftType == "right" {
			shift = asposecells.ShiftType_Right
		}
		return cells.DeleteRange(int32(beginRow), int32(beginColumn), int32(beginRow+rows-1), int32(beginColumn+columns-1), shift)
	}
}

// DeleteColumns creates a WorksheetAction that deletes a specified number of columns
// starting from the given column index.
//
// Parameters:
//   - beginColumn: The zero-based index of the first column to delete.
//   - columns: The total number of columns to delete.
//   - updateReference: A boolean indicating whether references (e.g., formulas, named ranges)
//     in other parts of the workbook should be updated to reflect the deletion.
//
// Returns:
//   - WorksheetAction: A function that deletes the specified columns from the worksheet.
func DeleteColumns(beginColumn int, columns int, updateReference bool) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, err := worksheet.GetCells()
		if err != nil {
			return err
		}
		return cells.DeleteColumns_Int_Int_Bool(int32(beginColumn), int32(columns), updateReference)
	}
}

// DeleteBlankColumns creates a WorksheetAction that removes all completely blank columns
// from the worksheet. A column is considered blank if none of its cells contain data or formatting.
//
// Returns:
//   - WorksheetAction: A function that modifies the worksheet by deleting empty columns.
func DeleteBlankColumns() WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, err := worksheet.GetCells()
		if err != nil {
			return err
		}
		return cells.DeleteBlankColumns()
	}
}

package editor

import (
	"fmt"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"strings"
	"time"
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
		cells, _ := worksheet.GetCells()
		cell, _ := cells.Get_Int_Int(int32(row), int32(column))
		switch v := value.(type) {
		case int8:
			obj, _ := asposecells.NewObject_Integer8(v)
			cell.PutValue_Object(obj)
			break
		case uint16:
			obj, _ := asposecells.NewObject_UInteger16(v)
			cell.PutValue_Object(obj)
			break
		case uint64:
			obj, _ := asposecells.NewObject_ULong(v)
			cell.PutValue_Object(obj)
			break
		case int16:
			obj, _ := asposecells.NewObject_Int16(v)
			cell.PutValue_Object(obj)
			break
		case int32:
			cell.PutValue_Int(v)
			break
		case int:
		case int64:
			obj, _ := asposecells.NewObject_Int64(v)
			cell.PutValue_Object(obj)
			break
		case float32:
			obj, _ := asposecells.NewObject_Float(v)
			cell.PutValue_Object(obj)
			break
		case float64:
			obj, _ := asposecells.NewObject_Double(v)
			cell.PutValue_Object(obj)
			break
		case string:
			cell.PutValue_String(v)
			break
		case bool:
			cell.PutValue_Bool(v)
			break
		case time.Time:
			obj, _ := asposecells.NewObject_Date(v)
			cell.PutValue_Object(obj)
			break
		default:
			return fmt.Errorf("invalid value: %v", value)
			break
		}
		return nil
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
		cells, _ := worksheet.GetCells()
		cellsRange, _ := cells.CreateRange_Int_Int_Int_Int(int32(beginRow), int32(beginColumn), int32(rows), int32(columns))
		switch v := value.(type) {
		case int8:
			obj, _ := asposecells.NewObject_Integer8(v)
			cellsRange.SetValue(obj)
			break
		case uint16:
			obj, _ := asposecells.NewObject_UInteger16(v)
			cellsRange.SetValue(obj)
			break
		case uint64:
			obj, _ := asposecells.NewObject_ULong(v)
			cellsRange.SetValue(obj)
			break
		case int16:
			obj, _ := asposecells.NewObject_Int16(v)
			cellsRange.SetValue(obj)
			break
		case int32:
			obj, _ := asposecells.NewObject_Int(v)
			cellsRange.SetValue(obj)
			break
		case int:
		case int64:
			obj, _ := asposecells.NewObject_Int64(v)
			cellsRange.SetValue(obj)
			break
		case float32:
			obj, _ := asposecells.NewObject_Float(v)
			cellsRange.SetValue(obj)
			break
		case float64:
			obj, _ := asposecells.NewObject_Double(v)
			cellsRange.SetValue(obj)
			break
		case string:
			obj, _ := asposecells.NewObject_String(v)
			cellsRange.SetValue(obj)
			break
		case bool:
			obj, _ := asposecells.NewObject_Bool(v)
			cellsRange.SetValue(obj)
			break
		case time.Time:
			obj, _ := asposecells.NewObject_Date(v)
			cellsRange.SetValue(obj)
			break
		default:
			return fmt.Errorf("invalid value: %v", value)
			break
		}
		return nil
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
		cells, _ := worksheet.GetCells()
		cells.Merge_Int_Int_Int_Int_Bool(int32(beginRow), int32(beginColumn), int32(rows), int32(columns), mergeConflict)
		return nil
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
		cells, _ := worksheet.GetCells()
		cells.UnMerge(int32(beginRow), int32(beginColumn), int32(rows), int32(columns))
		return nil
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
		cells, _ := worksheet.GetCells()
		cells.InsertRows_Int_Int_Bool(int32(beginRow), int32(rows), updateReference)
		return nil
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
		cells, _ := worksheet.GetCells()
		cells.InsertColumns_Int_Int_Bool(int32(beginColumn), int32(columns), updateReference)
		return nil
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
		cells, _ := worksheet.GetCells()
		cellsRange, _ := cells.CreateRange_Int_Int_Int_Int(int32(beginRow), int32(beginColumn), int32(rows), int32(columns))
		cellsRange.ClearContents()
		return nil
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
		cells, _ := worksheet.GetCells()
		cellsRange, _ := cells.CreateRange_Int_Int_Int_Int(int32(beginRow), int32(beginColumn), int32(rows), int32(columns))
		cellsRange.ClearFormats()
		return nil
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
		cells, _ := worksheet.GetCells()
		cells.DeleteRows_Int_Int_Bool(int32(beginRow), int32(rows), updateReference)
		return nil
	}
}

// DeleteBlankRows creates a WorksheetAction that removes all completely blank rows
// from the worksheet. A row is considered blank if it contains no data or formatting.
//
// Returns:
//   - WorksheetAction: A function that modifies the worksheet by deleting empty rows.
func DeleteBlankRows() WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, _ := worksheet.GetCells()
		cells.DeleteBlankRows()
		return nil
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
		cells, _ := worksheet.GetCells()
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
		cells.DeleteRange(int32(beginRow), int32(beginColumn), int32(beginRow+rows-1), int32(beginColumn+columns-1), shift)
		return nil
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
		cells, _ := worksheet.GetCells()
		cells.DeleteColumns_Int_Int_Bool(int32(beginColumn), int32(columns), updateReference)
		return nil
	}
}

// DeleteBlankColumns creates a WorksheetAction that removes all completely blank columns
// from the worksheet. A column is considered blank if none of its cells contain data or formatting.
//
// Returns:
//   - WorksheetAction: A function that modifies the worksheet by deleting empty columns.
func DeleteBlankColumns() WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, _ := worksheet.GetCells()
		cells.DeleteBlankColumns()
		return nil
	}
}

package editor

import (
	"fmt"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"strings"
	"time"
)

func SetCellValue(row, col int, value interface{}) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, _ := worksheet.GetCells()
		cell, _ := cells.Get_Int_Int(int32(row), int32(col))
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

func SetRangeValue(firstRow, firstColumn, totalRows, totalColumns int, value interface{}) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, _ := worksheet.GetCells()
		cellsRange, _ := cells.CreateRange_Int_Int_Int_Int(int32(firstRow), int32(firstColumn), int32(totalRows), int32(totalColumns))
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

func Merge(startRow, startColumn, rows, columns int, mergeConflict bool) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, _ := worksheet.GetCells()
		cells.Merge_Int_Int_Int_Int_Bool(int32(startRow), int32(startColumn), int32(rows), int32(columns), mergeConflict)
		return nil
	}
}

func UnMerge(startRow, startColumn, rows, columns int) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, _ := worksheet.GetCells()
		cells.UnMerge(int32(startRow), int32(startColumn), int32(rows), int32(columns))
		return nil
	}
}
func InsertRows(beginRow int, rows int, updateReference bool) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, _ := worksheet.GetCells()
		cells.InsertRows_Int_Int_Bool(int32(beginRow), int32(rows), updateReference)
		return nil
	}
}

func InsertColumns(beginColumn int, columns int, updateReference bool) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, _ := worksheet.GetCells()
		cells.InsertColumns_Int_Int_Bool(int32(beginColumn), int32(columns), updateReference)
		return nil
	}
}

func ClearRangeContents(_range string) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, _ := worksheet.GetCells()
		cellsRange, _ := cells.CreateRange_String(_range)
		cellsRange.ClearContents()
		return nil
	}
}

func ClearRangeFormats(_range string) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, _ := worksheet.GetCells()
		cellsRange, _ := cells.CreateRange_String(_range)
		cellsRange.ClearFormats()
		return nil
	}
}

func DeleteRows(beginRow int, rows int, updateReference bool) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, _ := worksheet.GetCells()
		cells.DeleteRows_Int_Int_Bool(int32(beginRow), int32(rows), updateReference)
		return nil
	}
}

func DeleteBlankRows() WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, _ := worksheet.GetCells()
		cells.DeleteBlankRows()
		return nil
	}
}

func DeleteRange(beginRow, beginColumn, endRow, endColumn int, shiftType string) WorksheetAction {
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
		cells.DeleteRange(int32(beginRow), int32(beginColumn), int32(endRow), int32(endColumn), shift)
		return nil
	}
}

func DeleteColumns(beginColumn int, columns int, updateReference bool) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, _ := worksheet.GetCells()
		cells.DeleteColumns_Int_Int_Bool(int32(beginColumn), int32(columns), updateReference)
		return nil
	}
}

func DeleteBlankColumns() WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		cells, _ := worksheet.GetCells()
		cells.DeleteBlankColumns()
		return nil
	}
}

package query

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// ConditionalFormattingInfo describes a conditional formatting collection's
// metadata: its areas and conditions. It is the read counterpart of
// editor.AddConditionalFormatting and editor.InConditionalFormatting.
type ConditionalFormattingInfo struct {
	// Index is the conditional formatting's zero-based position in the
	// worksheet's conditional formatting collection.
	Index int

	// Areas is the list of cell ranges the conditional formatting applies to,
	// in A1 notation (e.g., ["A1:A10", "C1:C10"]).
	Areas []string

	// Conditions is the list of conditions in this conditional formatting.
	Conditions []ConditionInfo
}

// ConditionInfo describes a single conditional formatting condition.
type ConditionInfo struct {
	// Index is the condition's zero-based position in the conditional
	// formatting's condition collection.
	Index int

	// Type is the condition type as a toolkit-native string (e.g., "colorScale",
	// "dataBar", "iconSet", "cellValue"). See editor.FormatConditionType for the
	// accepted names.
	Type string

	// Operator is the comparison operator as a toolkit-native string (e.g.,
	// "between", "greaterThan"). See editor.OperatorType for the accepted names.
	Operator string

	// Formula1 is the first formula or value for the condition.
	Formula1 string

	// Formula2 is the second formula or value for the condition (for "between"
	// operator).
	Formula2 string
}

// ConditionalFormattingCount returns the number of conditional formatting
// collections on the specified worksheet.
func ConditionalFormattingCount(src datasource.DataSource, opts ...Option) (int, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if src == nil {
		return 0, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(src)
	if err != nil {
		return 0, err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return 0, err
	}
	formattings, err := ws.GetConditionalFormattings()
	if err != nil {
		return 0, err
	}
	count, err := formattings.GetCount()
	if err != nil {
		return 0, err
	}
	return int(count), nil
}

// ConditionalFormattingInfoAt returns metadata about the conditional formatting
// collection at the given index on the specified worksheet.
func ConditionalFormattingInfoAt(src datasource.DataSource, index int, opts ...Option) (ConditionalFormattingInfo, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if src == nil {
		return ConditionalFormattingInfo{}, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(src)
	if err != nil {
		return ConditionalFormattingInfo{}, err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return ConditionalFormattingInfo{}, err
	}
	formattings, err := ws.GetConditionalFormattings()
	if err != nil {
		return ConditionalFormattingInfo{}, err
	}
	count, err := formattings.GetCount()
	if err != nil {
		return ConditionalFormattingInfo{}, err
	}
	if index < 0 || int32(index) >= count {
		return ConditionalFormattingInfo{}, toolkiterrors.ErrConditionNotFound
	}
	collection, err := formattings.Get(int32(index))
	if err != nil {
		return ConditionalFormattingInfo{}, err
	}
	return readConditionalFormattingInfo(collection, index)
}

// AllConditionalFormattings returns metadata about all conditional formatting
// collections on the specified worksheet.
func AllConditionalFormattings(src datasource.DataSource, opts ...Option) ([]ConditionalFormattingInfo, error) {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if src == nil {
		return nil, toolkiterrors.ErrDataSourceNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(src)
	if err != nil {
		return nil, err
	}
	ws, err := sheetFor(cfg, workbook)
	if err != nil {
		return nil, err
	}
	formattings, err := ws.GetConditionalFormattings()
	if err != nil {
		return nil, err
	}
	count, err := formattings.GetCount()
	if err != nil {
		return nil, err
	}
	result := make([]ConditionalFormattingInfo, 0, count)
	for i := 0; i < int(count); i++ {
		collection, err := formattings.Get(int32(i))
		if err != nil {
			return nil, err
		}
		info, err := readConditionalFormattingInfo(collection, i)
		if err != nil {
			return nil, err
		}
		result = append(result, info)
	}
	return result, nil
}

// readConditionalFormattingInfo extracts metadata from a conditional formatting collection.
func readConditionalFormattingInfo(collection *asposecells.FormatConditionCollection, index int) (ConditionalFormattingInfo, error) {
	info := ConditionalFormattingInfo{Index: index}

	// Areas
	rangeCount, err := collection.GetRangeCount()
	if err != nil {
		return ConditionalFormattingInfo{}, err
	}
	info.Areas = make([]string, 0, rangeCount)
	for i := int32(0); i < rangeCount; i++ {
		area, err := collection.GetCellArea(i)
		if err != nil {
			return ConditionalFormattingInfo{}, err
		}
		s, err := area.ToString()
		if err != nil {
			return ConditionalFormattingInfo{}, err
		}
		info.Areas = append(info.Areas, s)
	}

	// Conditions
	condCount, err := collection.GetCount()
	if err != nil {
		return ConditionalFormattingInfo{}, err
	}
	info.Conditions = make([]ConditionInfo, 0, condCount)
	for i := 0; i < int(condCount); i++ {
		cond, err := collection.Get(int32(i))
		if err != nil {
			return ConditionalFormattingInfo{}, err
		}
		condInfo, err := readConditionInfo(cond, i)
		if err != nil {
			return ConditionalFormattingInfo{}, err
		}
		info.Conditions = append(info.Conditions, condInfo)
	}

	return info, nil
}

// readConditionInfo extracts metadata from a condition.
func readConditionInfo(cond *asposecells.FormatCondition, index int) (ConditionInfo, error) {
	info := ConditionInfo{Index: index}

	// Type
	ct, err := cond.GetType()
	if err != nil {
		return ConditionInfo{}, err
	}
	info.Type = formatConditionTypeName(int32(ct))

	// Operator
	op, err := cond.GetOperator()
	if err != nil {
		return ConditionInfo{}, err
	}
	info.Operator = operatorTypeName(int32(op))

	// Formulas
	info.Formula1, err = cond.GetFormula1()
	if err != nil {
		return ConditionInfo{}, err
	}
	info.Formula2, err = cond.GetFormula2()
	if err != nil {
		return ConditionInfo{}, err
	}

	return info, nil
}

// formatConditionTypeName maps format condition type enum values to toolkit-native names.
func formatConditionTypeName(ct int32) string {
	switch ct {
	case 1:
		return "cellValue"
	case 2:
		return "expression"
	case 4:
		return "top10"
	case 8:
		return "uniqueValues"
	case 16:
		return "duplicateValues"
	case 32:
		return "containsText"
	case 64:
		return "notContainsText"
	case 128:
		return "beginsWith"
	case 256:
		return "endsWith"
	case 512:
		return "containsBlanks"
	case 1024:
		return "notContainsBlanks"
	case 2048:
		return "containsErrors"
	case 4096:
		return "notContainsErrors"
	case 8192:
		return "timePeriod"
	case 16384:
		return "aboveAverage"
	case 32768:
		return "colorScale"
	case 65536:
		return "dataBar"
	case 131072:
		return "iconSet"
	default:
		return "unknown"
	}
}

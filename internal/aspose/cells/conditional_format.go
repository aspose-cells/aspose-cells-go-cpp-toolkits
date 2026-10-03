package cells

import (
	"fmt"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// formatConditionTypeByName maps normalized condition type names to engine enums.
var formatConditionTypeByName = map[string]asposecells.FormatConditionType{
	"cellvalue":         asposecells.FormatConditionType_CellValue,
	"expression":        asposecells.FormatConditionType_Expression,
	"top10":             asposecells.FormatConditionType_Top10,
	"uniquevalues":      asposecells.FormatConditionType_UniqueValues,
	"duplicatevalues":   asposecells.FormatConditionType_DuplicateValues,
	"containstext":      asposecells.FormatConditionType_ContainsText,
	"notcontainstext":   asposecells.FormatConditionType_NotContainsText,
	"beginswith":        asposecells.FormatConditionType_BeginsWith,
	"endswith":          asposecells.FormatConditionType_EndsWith,
	"containsblanks":    asposecells.FormatConditionType_ContainsBlanks,
	"notcontainsblanks": asposecells.FormatConditionType_NotContainsBlanks,
	"containserrors":    asposecells.FormatConditionType_ContainsErrors,
	"notcontainserrors": asposecells.FormatConditionType_NotContainsErrors,
	"timeperiod":        asposecells.FormatConditionType_TimePeriod,
	"aboveaverage":      asposecells.FormatConditionType_AboveAverage,
	"colorscale":        asposecells.FormatConditionType_ColorScale,
	"databar":           asposecells.FormatConditionType_DataBar,
	"iconset":           asposecells.FormatConditionType_IconSet,
}

// ResolveFormatConditionType normalizes a string name and looks up the engine enum.
func ResolveFormatConditionType(name string) (asposecells.FormatConditionType, error) {
	key := normalizeEnumName(name)
	if ct, ok := formatConditionTypeByName[key]; ok {
		return ct, nil
	}
	return 0, fmt.Errorf("format condition type %q: %w", name, toolkiterrors.ErrInvalidFormatConditionType)
}

// iconSetTypeByName maps normalized icon set type names to engine enums.
var iconSetTypeByName = map[string]asposecells.IconSetType{
	"arrows3":         asposecells.IconSetType_Arrows3,
	"arrowsgray3":     asposecells.IconSetType_ArrowsGray3,
	"flags3":          asposecells.IconSetType_Flags3,
	"signs3":          asposecells.IconSetType_Signs3,
	"symbols3":        asposecells.IconSetType_Symbols3,
	"symbols32":       asposecells.IconSetType_Symbols32,
	"trafficlights31": asposecells.IconSetType_TrafficLights31,
	"trafficlights32": asposecells.IconSetType_TrafficLights32,
	"arrows4":         asposecells.IconSetType_Arrows4,
	"arrowsgray4":     asposecells.IconSetType_ArrowsGray4,
	"rating4":         asposecells.IconSetType_Rating4,
	"redtoblack4":     asposecells.IconSetType_RedToBlack4,
	"trafficlights4":  asposecells.IconSetType_TrafficLights4,
	"arrows5":         asposecells.IconSetType_Arrows5,
	"arrowsgray5":     asposecells.IconSetType_ArrowsGray5,
	"quarters5":       asposecells.IconSetType_Quarters5,
	"rating5":         asposecells.IconSetType_Rating5,
	"stars3":          asposecells.IconSetType_Stars3,
	"boxes5":          asposecells.IconSetType_Boxes5,
	"triangles3":      asposecells.IconSetType_Triangles3,
}

// ResolveIconSetType normalizes a string name and looks up the engine enum.
func ResolveIconSetType(name string) (asposecells.IconSetType, error) {
	key := normalizeEnumName(name)
	if is, ok := iconSetTypeByName[key]; ok {
		return is, nil
	}
	return 0, fmt.Errorf("icon set type %q: %w", name, toolkiterrors.ErrInvalidIconSetType)
}

// AddConditionalFormatting creates a conditional formatting collection for the
// given cell area on the worksheet. The area string is parsed as an A1-style
// range (e.g. "A1:A10").
func AddConditionalFormatting(ws *asposecells.Worksheet, areaStr string) (*asposecells.FormatConditionCollection, error) {
	area, err := ParseArea(areaStr)
	if err != nil {
		return nil, err
	}
	if err := ValidateGridArea(area); err != nil {
		return nil, err
	}
	formattings, err := engine.Derive(ws.GetConditionalFormattings())
	if err != nil {
		return nil, err
	}
	// Add creates a new format condition collection and returns its index
	idx, err := formattings.Add()
	if err != nil {
		return nil, err
	}
	// Add an area to the collection
	collection, err := formattings.Get(idx)
	if err != nil {
		return nil, err
	}
	cellArea, err := asposecells.CellArea_CreateCellArea_Int_Int_Int_Int(
		int32(area.Start.Row), int32(area.Start.Col),
		int32(area.End.Row), int32(area.End.Col),
	)
	if err != nil {
		return nil, err
	}
	_, err = collection.AddArea(cellArea)
	if err != nil {
		return nil, err
	}
	return collection, nil
}

// AddConditionToCollection adds a condition to an existing FormatConditionCollection.
func AddConditionToCollection(collection *asposecells.FormatConditionCollection, condType asposecells.FormatConditionType, op asposecells.OperatorType, formula1, formula2 string) (int, error) {
	idx, err := collection.AddCondition_FormatConditionType_OperatorType_String_String(
		condType, op, formula1, formula2,
	)
	if err != nil {
		return 0, err
	}
	return int(idx), nil
}

// GetCondition returns the condition at the given index from the collection.
func GetCondition(collection *asposecells.FormatConditionCollection, index int) (*asposecells.FormatCondition, error) {
	count, err := collection.GetCount()
	if err != nil {
		return nil, err
	}
	if index < 0 || int32(index) >= count {
		return nil, fmt.Errorf("condition index %d out of range [0, %d): %w",
			index, count, toolkiterrors.ErrConditionNotFound)
	}
	return collection.Get(int32(index))
}

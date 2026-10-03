package cells

import (
	"fmt"
	"strings"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// normalizeEnumName lowercases and strips non-alphanumeric characters, so
// "Whole Number", "whole_number", "WholeNumber" all resolve the same way.
func normalizeEnumName(name string) string {
	var b strings.Builder
	b.Grow(len(name))
	for _, r := range name {
		switch {
		case r >= 'a' && r <= 'z', r >= '0' && r <= '9':
			b.WriteRune(r)
		case r >= 'A' && r <= 'Z':
			b.WriteRune(r + ('a' - 'A'))
		}
	}
	return b.String()
}

// validationTypeByName maps normalized validation type names to engine enums.
var validationTypeByName = map[string]asposecells.ValidationType{
	"anyvalue":    asposecells.ValidationType_AnyValue,
	"wholenumber": asposecells.ValidationType_WholeNumber,
	"integer":     asposecells.ValidationType_WholeNumber,
	"decimal":     asposecells.ValidationType_Decimal,
	"list":        asposecells.ValidationType_List,
	"date":        asposecells.ValidationType_Date,
	"time":        asposecells.ValidationType_Time,
	"textlength":  asposecells.ValidationType_TextLength,
	"custom":      asposecells.ValidationType_Custom,
}

// ResolveValidationType normalizes a string name and looks up the engine enum.
// Returns ErrInvalidValidationType for unknown names.
func ResolveValidationType(name string) (asposecells.ValidationType, error) {
	key := normalizeEnumName(name)
	if vt, ok := validationTypeByName[key]; ok {
		return vt, nil
	}
	return 0, fmt.Errorf("validation type %q: %w", name, toolkiterrors.ErrInvalidValidationType)
}

// operatorTypeByName maps normalized operator names to engine enums. Keep this
// in step with query.operatorTypeName, which maps the ordinals back to names.
var operatorTypeByName = map[string]asposecells.OperatorType{
	"between":        asposecells.OperatorType_Between,
	"equal":          asposecells.OperatorType_Equal,
	"notequal":       asposecells.OperatorType_NotEqual,
	"notbetween":     asposecells.OperatorType_NotBetween,
	"none":           asposecells.OperatorType_None,
	"lessthan":       asposecells.OperatorType_LessThan,
	"lessorequal":    asposecells.OperatorType_LessOrEqual,
	"greaterthan":    asposecells.OperatorType_GreaterThan,
	"greaterorequal": asposecells.OperatorType_GreaterOrEqual,
}

// ResolveOperatorType normalizes a string name and looks up the engine enum.
// Returns ErrInvalidOperatorType for unknown names.
func ResolveOperatorType(name string) (asposecells.OperatorType, error) {
	key := normalizeEnumName(name)
	if op, ok := operatorTypeByName[key]; ok {
		return op, nil
	}
	return 0, fmt.Errorf("operator type %q: %w", name, toolkiterrors.ErrInvalidOperatorType)
}

// validationAlertByName maps normalized alert type names to engine enums.
var validationAlertByName = map[string]asposecells.ValidationAlertType{
	"information": asposecells.ValidationAlertType_Information,
	"warning":     asposecells.ValidationAlertType_Warning,
	"stop":        asposecells.ValidationAlertType_Stop,
}

// ResolveValidationAlertType normalizes a string name and looks up the engine enum.
func ResolveValidationAlertType(name string) (asposecells.ValidationAlertType, error) {
	key := normalizeEnumName(name)
	if at, ok := validationAlertByName[key]; ok {
		return at, nil
	}
	return 0, fmt.Errorf("validation alert type %q: %w", name, toolkiterrors.ErrInvalidValidationType)
}

// AddValidation adds a data validation to the worksheet for the given cell area.
// The area string is parsed as an A1-style range (e.g. "A1:A10").
func AddValidation(ws *asposecells.Worksheet, areaStr string) (*asposecells.Validation, error) {
	area, err := ParseArea(areaStr)
	if err != nil {
		return nil, err
	}
	if err := ValidateGridArea(area); err != nil {
		return nil, err
	}
	validations, err := engine.Derive(ws.GetValidations())
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
	idx, err := validations.Add(cellArea)
	if err != nil {
		return nil, err
	}
	return validations.Get(idx)
}

// GetValidation returns the validation at the given index, or ErrValidationNotFound.
func GetValidation(validations *asposecells.ValidationCollection, index int) (*asposecells.Validation, error) {
	count, err := validations.GetCount()
	if err != nil {
		return nil, err
	}
	if index < 0 || int32(index) >= count {
		return nil, fmt.Errorf("validation index %d out of range [0, %d): %w",
			index, count, toolkiterrors.ErrValidationNotFound)
	}
	return validations.Get(int32(index))
}

package editor

import (
	"fmt"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// FormatConditionType names a conditional formatting rule type. The constants
// below cover the common rule types; any of the engine's condition types is
// also accepted by its own name (for example "colorScale" or "dataBar"), and
// names are matched case- and punctuation-insensitively.
type FormatConditionType string

// The curated condition types. Each names a category of conditional formatting rule.
const (
	FormatConditionTypeCellValue    FormatConditionType = "cellValue"
	FormatConditionTypeExpression   FormatConditionType = "expression"
	FormatConditionTypeTop10        FormatConditionType = "top10"
	FormatConditionTypeUniqueValues FormatConditionType = "uniqueValues"
	FormatConditionTypeDuplicate    FormatConditionType = "duplicateValues"
	FormatConditionTypeContainsText FormatConditionType = "containsText"
	FormatConditionTypeAboveAverage FormatConditionType = "aboveAverage"
	FormatConditionTypeColorScale   FormatConditionType = "colorScale"
	FormatConditionTypeDataBar      FormatConditionType = "dataBar"
	FormatConditionTypeIconSet      FormatConditionType = "iconSet"
)

// IconSetType names an icon set for icon set conditional formatting. The
// constants below cover the common icon sets; any of the engine's icon set
// types is also accepted by its own name (for example "trafficLights3" or
// "rating5"), and names are matched case- and punctuation-insensitively.
type IconSetType string

// The curated icon set types. Each names a visual icon set.
const (
	IconSetArrows3         IconSetType = "arrows3"
	IconSetArrowsGray3     IconSetType = "arrowsGray3"
	IconSetFlags3          IconSetType = "flags3"
	IconSetSigns3          IconSetType = "signs3"
	IconSetSymbols3        IconSetType = "symbols3"
	IconSetTrafficLights31 IconSetType = "trafficLights31"
	IconSetTrafficLights32 IconSetType = "trafficLights32"
	IconSetArrows4         IconSetType = "arrows4"
	IconSetRating4         IconSetType = "rating4"
	IconSetArrows5         IconSetType = "arrows5"
	IconSetRating5         IconSetType = "rating5"
	IconSetStars3          IconSetType = "stars3"
)

// AddConditionalFormatting creates a WorkbookAction that adds a conditional
// formatting rule to the specified cell range on the given worksheet. The range
// is an A1-style reference (e.g. "A1:A10"). The conditional formatting is
// configured by the provided ConditionalFormatActions.
//
// Parameters:
//   - sheetID: The worksheet identifier (name or index).
//   - range: The cell range to format (e.g. "A1:A10", "B2:B100").
//   - actions: A variadic list of ConditionalFormatAction functions to configure
//     the conditional formatting rules.
//
// Returns:
//   - WorkbookAction: A function that adds the conditional formatting to the worksheet.
//
// Example:
//
//	editor.AddConditionalFormatting("Sheet1", "A1:A10",
//	    editor.WithColorScale("#F8696B", "#FFFFFF", "#63BE7B"),
//	)
func AddConditionalFormatting(sheetID interface{}, cellRange string, actions ...ConditionalFormatAction) WorkbookAction {
	return func(wb *asposecells.Workbook) error {
		sheet, err := resolveWorksheet(wb, sheetID)
		if err != nil {
			return err
		}
		collection, err := cells.AddConditionalFormatting(sheet, cellRange)
		if err != nil {
			return err
		}
		for _, action := range actions {
			if err := action(collection); err != nil {
				return err
			}
		}
		return nil
	}
}

// InConditionalFormatting creates a WorksheetAction that applies conditional
// formatting actions to the worksheet's existing conditional formatting collection
// at the given index. Collections are indexed from zero in the order they were added.
//
// Parameters:
//   - index: The zero-based index of the conditional formatting collection to modify.
//   - actions: A variadic list of ConditionalFormatAction functions to apply.
//
// Returns:
//   - WorksheetAction: A function that modifies the conditional formatting.
//
// Example:
//
//	editor.InWorksheet("Sheet1",
//	    editor.InConditionalFormatting(0,
//	        editor.WithDataBar("#63BE7B"),
//	    ),
//	)
func InConditionalFormatting(index int, actions ...ConditionalFormatAction) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		formattings, err := worksheet.GetConditionalFormattings()
		if err != nil {
			return err
		}
		count, err := formattings.GetCount()
		if err != nil {
			return err
		}
		if index < 0 || int32(index) >= count {
			return fmt.Errorf("conditional formatting index %d out of range [0, %d): %w",
				index, count, toolkiterrors.ErrConditionNotFound)
		}
		collection, err := formattings.Get(int32(index))
		if err != nil {
			return err
		}
		for _, action := range actions {
			if err := action(collection); err != nil {
				return err
			}
		}
		return nil
	}
}

// DeleteConditionalFormatting creates a WorksheetAction that removes the
// conditional formatting collection at the given index from the worksheet.
// Note: This removes the entire conditional formatting rule including all its
// conditions and areas.
//
// Parameters:
//   - index: The zero-based index of the conditional formatting to delete.
//
// Returns:
//   - WorksheetAction: A function that deletes the conditional formatting.
func DeleteConditionalFormatting(index int) WorksheetAction {
	return func(worksheet *asposecells.Worksheet) error {
		formattings, err := worksheet.GetConditionalFormattings()
		if err != nil {
			return err
		}
		count, err := formattings.GetCount()
		if err != nil {
			return err
		}
		if index < 0 || int32(index) >= count {
			return fmt.Errorf("conditional formatting index %d out of range [0, %d): %w",
				index, count, toolkiterrors.ErrConditionNotFound)
		}
		// Get the collection to find its areas
		collection, err := formattings.Get(int32(index))
		if err != nil {
			return err
		}
		// Remove all conditions from the collection first
		condCount, err := collection.GetCount()
		if err != nil {
			return err
		}
		for i := int32(condCount) - 1; i >= 0; i-- {
			if err := collection.RemoveCondition(i); err != nil {
				return err
			}
		}
		// Remove all areas
		areaCount, err := collection.GetRangeCount()
		if err != nil {
			return err
		}
		for i := int32(areaCount) - 1; i >= 0; i-- {
			area, err := collection.GetCellArea(i)
			if err != nil {
				return err
			}
			startRow, err := area.Get_StartRow()
			if err != nil {
				return err
			}
			startCol, err := area.Get_StartColumn()
			if err != nil {
				return err
			}
			endRow, err := area.Get_EndRow()
			if err != nil {
				return err
			}
			endCol, err := area.Get_EndColumn()
			if err != nil {
				return err
			}
			if err := formattings.RemoveArea(
				startRow, startCol,
				endRow-startRow+1,
				endCol-startCol+1,
			); err != nil {
				return err
			}
		}
		return nil
	}
}

// WithColorScale creates a ConditionalFormatAction that adds a 2-color or
// 3-color scale conditional formatting rule. A color scale applies a gradient
// fill to cells based on their values.
//
// Parameters:
//   - minColor: The color for the minimum value (e.g. "#F8696B" for red).
//   - maxColor: The color for the maximum value (e.g. "#63BE7B" for green).
//   - midColor: Optional middle color for a 3-color scale (e.g. "#FFFFFF" for white).
//     Pass nil or empty string to use a 2-color scale.
//
// Returns:
//   - ConditionalFormatAction: A function that adds a color scale rule.
//
// Example:
//
//	editor.WithColorScale("#F8696B", "#63BE7B", nil)           // 2-color scale
//	editor.WithColorScale("#F8696B", "#63BE7B", "#FFFFFF")     // 3-color scale
func WithColorScale(minColor, maxColor interface{}, midColor interface{}) ConditionalFormatAction {
	return func(collection *asposecells.FormatConditionCollection) error {
		// Add a color scale condition
		idx, err := cells.AddConditionToCollection(collection,
			asposecells.FormatConditionType_ColorScale,
			asposecells.OperatorType_None,
			"", "",
		)
		if err != nil {
			return err
		}
		cond, err := cells.GetCondition(collection, idx)
		if err != nil {
			return err
		}
		cs, err := cond.GetColorScale()
		if err != nil {
			return err
		}
		// Set min color
		minC, err := resolveColor(minColor)
		if err != nil {
			return err
		}
		if err := cs.SetMinColor(minC); err != nil {
			return err
		}
		// Set max color
		maxC, err := resolveColor(maxColor)
		if err != nil {
			return err
		}
		if err := cs.SetMaxColor(maxC); err != nil {
			return err
		}
		// Set mid color if provided (3-color scale)
		if midColor != nil {
			if midStr, ok := midColor.(string); ok && midStr != "" {
				midC, err := resolveColor(midColor)
				if err != nil {
					return err
				}
				if err := cs.SetIs3ColorScale(true); err != nil {
					return err
				}
				return cs.SetMidColor(midC)
			}
		}
		return nil
	}
}

// WithDataBar creates a ConditionalFormatAction that adds a data bar conditional
// formatting rule. A data bar displays a horizontal bar in each cell proportional
// to the cell's value.
//
// Parameters:
//   - color: The color of the data bar (e.g. "#63BE7B" for green).
//
// Returns:
//   - ConditionalFormatAction: A function that adds a data bar rule.
//
// Example:
//
//	editor.WithDataBar("#63BE7B")
func WithDataBar(color interface{}) ConditionalFormatAction {
	return func(collection *asposecells.FormatConditionCollection) error {
		idx, err := cells.AddConditionToCollection(collection,
			asposecells.FormatConditionType_DataBar,
			asposecells.OperatorType_None,
			"", "",
		)
		if err != nil {
			return err
		}
		cond, err := cells.GetCondition(collection, idx)
		if err != nil {
			return err
		}
		db, err := cond.GetDataBar()
		if err != nil {
			return err
		}
		c, err := resolveColor(color)
		if err != nil {
			return err
		}
		return db.SetColor(c)
	}
}

// WithIconSet creates a ConditionalFormatAction that adds an icon set conditional
// formatting rule. An icon set displays icons in cells based on their values
// relative to thresholds.
//
// Parameters:
//   - iconSetType: The type of icon set (e.g. IconSetTrafficLights31, IconSetArrows3).
//
// Returns:
//   - ConditionalFormatAction: A function that adds an icon set rule.
//
// Example:
//
//	editor.WithIconSet(editor.IconSetTrafficLights31)
func WithIconSet(iconSetType IconSetType) ConditionalFormatAction {
	return func(collection *asposecells.FormatConditionCollection) error {
		idx, err := cells.AddConditionToCollection(collection,
			asposecells.FormatConditionType_IconSet,
			asposecells.OperatorType_None,
			"", "",
		)
		if err != nil {
			return err
		}
		cond, err := cells.GetCondition(collection, idx)
		if err != nil {
			return err
		}
		is, err := cond.GetIconSet()
		if err != nil {
			return err
		}
		ist, err := cells.ResolveIconSetType(string(iconSetType))
		if err != nil {
			return err
		}
		return is.SetType(ist)
	}
}

// WithCellValueRule creates a ConditionalFormatAction that adds a cell value
// comparison conditional formatting rule. The rule applies formatting when the
// cell value meets the specified condition.
//
// Parameters:
//   - operator: The comparison operator (e.g. OperatorTypeGreaterThan,
//     OperatorTypeBetween).
//   - formula1: The first formula or value for comparison.
//   - formula2: The second formula or value (for "between" operator). Pass empty
//     string for single-value operators.
//   - styleActions: A variadic list of StyleAction functions to apply when the
//     condition is met.
//
// Returns:
//   - ConditionalFormatAction: A function that adds a cell value rule.
//
// Example:
//
//	editor.WithCellValueRule(editor.OperatorTypeGreaterThan, "100", "",
//	    editor.WithFontColor("#FF0000"),
//	    editor.WithFontIsBold(true),
//	)
func WithCellValueRule(operator OperatorType, formula1, formula2 string, styleActions ...StyleAction) ConditionalFormatAction {
	return func(collection *asposecells.FormatConditionCollection) error {
		op, err := cells.ResolveOperatorType(string(operator))
		if err != nil {
			return err
		}
		idx, err := cells.AddConditionToCollection(collection,
			asposecells.FormatConditionType_CellValue,
			op,
			formula1, formula2,
		)
		if err != nil {
			return err
		}
		cond, err := cells.GetCondition(collection, idx)
		if err != nil {
			return err
		}
		// Apply style actions if provided
		if len(styleActions) > 0 {
			style, err := cond.GetStyle()
			if err != nil {
				return err
			}
			for _, action := range styleActions {
				if err := action(style); err != nil {
					return err
				}
			}
			return cond.SetStyle(style)
		}
		return nil
	}
}

// WithExpressionRule creates a ConditionalFormatAction that adds a formula-based
// conditional formatting rule. The rule applies formatting when the formula
// evaluates to true.
//
// Parameters:
//   - formula: The formula to evaluate (e.g. "=A1>100").
//   - styleActions: A variadic list of StyleAction functions to apply when the
//     condition is met.
//
// Returns:
//   - ConditionalFormatAction: A function that adds an expression rule.
//
// Example:
//
//	editor.WithExpressionRule("=A1>100",
//	    editor.WithBackgroundColor("#FFFF00"),
//	)
func WithExpressionRule(formula string, styleActions ...StyleAction) ConditionalFormatAction {
	return func(collection *asposecells.FormatConditionCollection) error {
		idx, err := cells.AddConditionToCollection(collection,
			asposecells.FormatConditionType_Expression,
			asposecells.OperatorType_None,
			formula, "",
		)
		if err != nil {
			return err
		}
		cond, err := cells.GetCondition(collection, idx)
		if err != nil {
			return err
		}
		// Apply style actions if provided
		if len(styleActions) > 0 {
			style, err := cond.GetStyle()
			if err != nil {
				return err
			}
			for _, action := range styleActions {
				if err := action(style); err != nil {
					return err
				}
			}
			return cond.SetStyle(style)
		}
		return nil
	}
}

// WithAboveAverageRule creates a ConditionalFormatAction that adds an above/below
// average conditional formatting rule.
//
// Parameters:
//   - styleActions: A variadic list of StyleAction functions to apply to cells
//     above average.
//
// Returns:
//   - ConditionalFormatAction: A function that adds an above average rule.
//
// Example:
//
//	editor.WithAboveAverageRule(
//	    editor.WithFontColor("#006100"),
//	    editor.WithBackgroundColor("#C6EFCE"),
//	)
func WithAboveAverageRule(styleActions ...StyleAction) ConditionalFormatAction {
	return func(collection *asposecells.FormatConditionCollection) error {
		idx, err := cells.AddConditionToCollection(collection,
			asposecells.FormatConditionType_AboveAverage,
			asposecells.OperatorType_None,
			"", "",
		)
		if err != nil {
			return err
		}
		cond, err := cells.GetCondition(collection, idx)
		if err != nil {
			return err
		}
		// Apply style actions if provided
		if len(styleActions) > 0 {
			style, err := cond.GetStyle()
			if err != nil {
				return err
			}
			for _, action := range styleActions {
				if err := action(style); err != nil {
					return err
				}
			}
			return cond.SetStyle(style)
		}
		return nil
	}
}

// WithTop10Rule creates a ConditionalFormatAction that adds a top/bottom N
// conditional formatting rule.
//
// Parameters:
//   - rank: The number of items to highlight (e.g. 10 for top 10).
//   - isTop: true for top N, false for bottom N.
//   - styleActions: A variadic list of StyleAction functions to apply.
//
// Returns:
//   - ConditionalFormatAction: A function that adds a top 10 rule.
//
// Example:
//
//	editor.WithTop10Rule(10, true,
//	    editor.WithFontColor("#9C5700"),
//	    editor.WithBackgroundColor("#FFEB9C"),
//	)
func WithTop10Rule(rank int, isTop bool, styleActions ...StyleAction) ConditionalFormatAction {
	return func(collection *asposecells.FormatConditionCollection) error {
		idx, err := cells.AddConditionToCollection(collection,
			asposecells.FormatConditionType_Top10,
			asposecells.OperatorType_None,
			"", "",
		)
		if err != nil {
			return err
		}
		cond, err := cells.GetCondition(collection, idx)
		if err != nil {
			return err
		}
		top10, err := cond.GetTop10()
		if err != nil {
			return err
		}
		if err := top10.SetRank(int32(rank)); err != nil {
			return err
		}
		// IsBottom is the inverse of isTop
		if err := top10.SetIsBottom(!isTop); err != nil {
			return err
		}
		// Apply style actions if provided
		if len(styleActions) > 0 {
			style, err := cond.GetStyle()
			if err != nil {
				return err
			}
			for _, action := range styleActions {
				if err := action(style); err != nil {
					return err
				}
			}
			return cond.SetStyle(style)
		}
		return nil
	}
}

// WithConditionalFormatRule creates a ConditionalFormatAction that adds a custom
// conditional formatting rule. This is a flexible function that allows you to
// specify any condition type, operator, and formulas. Use this when the higher-level
// convenience functions (WithCellValueRule, WithExpressionRule, etc.) don't meet
// your needs.
//
// Parameters:
//   - condType: The condition type (e.g., FormatConditionTypeCellValue,
//     FormatConditionTypeExpression).
//   - operator: The comparison operator (e.g., OperatorTypeGreaterThan,
//     OperatorTypeBetween). Use an empty string for types that don't need an operator.
//   - formula1: The first formula or value.
//   - formula2: The second formula or value (for operators like Between). Pass empty
//     string for single-value operators.
//   - styleActions: A variadic list of StyleAction functions to apply when the
//     condition is met.
//
// Returns:
//   - ConditionalFormatAction: A function that adds the custom rule.
//
// Example:
//
//	// Cell value greater than 100
//	editor.WithConditionalFormatRule(editor.FormatConditionTypeCellValue,
//	    editor.OperatorTypeGreaterThan, "100", "",
//	    editor.WithFontColor("#FF0000"),
//	)
//
//	// Expression-based rule
//	editor.WithConditionalFormatRule(editor.FormatConditionTypeExpression,
//	    "", "=AND(A1>0,B1<100)", "",
//	    editor.WithBackgroundColor("#FFFF00"),
//	)
//
//	// Duplicate values
//	editor.WithConditionalFormatRule(editor.FormatConditionTypeDuplicateValues,
//	    "", "", "",
//	    editor.WithFontColor("#9C0006"),
//	    editor.WithBackgroundColor("#FFC7CE"),
//	)
func WithConditionalFormatRule(condType FormatConditionType, operator OperatorType, formula1, formula2 string, styleActions ...StyleAction) ConditionalFormatAction {
	return func(collection *asposecells.FormatConditionCollection) error {
		ct, err := cells.ResolveFormatConditionType(string(condType))
		if err != nil {
			return err
		}
		op, err := cells.ResolveOperatorType(string(operator))
		if err != nil {
			// Some condition types don't need an operator, so use None as default
			op = asposecells.OperatorType_None
		}
		idx, err := cells.AddConditionToCollection(collection, ct, op, formula1, formula2)
		if err != nil {
			return err
		}
		cond, err := cells.GetCondition(collection, idx)
		if err != nil {
			return err
		}
		// Apply style actions if provided
		if len(styleActions) > 0 {
			style, err := cond.GetStyle()
			if err != nil {
				return err
			}
			for _, action := range styleActions {
				if err := action(style); err != nil {
					return err
				}
			}
			return cond.SetStyle(style)
		}
		return nil
	}
}

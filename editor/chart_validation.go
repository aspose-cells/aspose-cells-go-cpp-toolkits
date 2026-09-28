package editor

import (
	"fmt"

	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
)

// ValidateChartDataRange validates a chart data range before creating a chart.
// It checks that the range is well-formed, in-grid, and can produce a chart
// with meaningful data.
//
// This is a convenience wrapper around cells.ValidateChartDataRange that
// provides additional chart-specific guidance in error messages.
//
// Parameters:
//   - dataRange: The data range to validate (e.g., "A1:C5").
//   - byColumn: Whether the range will be read by column (true) or by row (false).
//
// Returns:
//   - error: An error if the range is invalid, with guidance on how to fix it.
//
// Example:
//
//	if err := editor.ValidateChartDataRange("A1:C5", true); err != nil {
//	    log.Printf("Invalid data range: %v", err)
//	}
func ValidateChartDataRange(dataRange string, byColumn bool) error {
	if err := cells.ValidateChartDataRange(dataRange); err != nil {
		return fmt.Errorf("chart data range %q is invalid: %w (hint: ensure the range is within the worksheet grid and contains multiple cells)", dataRange, err)
	}
	return nil
}

// ValidateChartBounds validates chart bounds before creating or moving a chart.
// It checks that the bounds are well-formed, in-grid, and enclose a positive
// area.
//
// This is a convenience wrapper around cells.ValidateChartBounds that provides
// additional guidance in error messages.
//
// Parameters:
//   - topRow: The zero-based row of the chart's top edge.
//   - leftColumn: The zero-based column of the chart's left edge.
//   - bottomRow: The zero-based row of the chart's bottom edge.
//   - rightColumn: The zero-based column of the chart's right edge.
//
// Returns:
//   - error: An error if the bounds are invalid, with guidance on how to fix them.
//
// Example:
//
//	if err := editor.ValidateChartBounds(5, 0, 20, 7); err != nil {
//	    log.Printf("Invalid chart bounds: %v", err)
//	}
func ValidateChartBounds(topRow, leftColumn, bottomRow, rightColumn int) error {
	if err := cells.ValidateChartBounds(topRow, leftColumn, bottomRow, rightColumn); err != nil {
		return fmt.Errorf("chart bounds (%d, %d, %d, %d) are invalid: %w (hint: bounds must be non-negative, non-reversed, and enclose a positive area within the worksheet grid)",
			topRow, leftColumn, bottomRow, rightColumn, err)
	}
	return nil
}

// ValidateChartStyle validates a chart style number. The engine's built-in
// styles are numbered 1..48; anything outside that range is rejected.
//
// Parameters:
//   - style: The style number to validate.
//
// Returns:
//   - error: An error if the style is outside 1..48.
//
// Example:
//
//	if err := editor.ValidateChartStyle(7); err != nil {
//	    log.Printf("Invalid chart style: %v", err)
//	}
func ValidateChartStyle(style int) error {
	if style < 1 || style > 48 {
		return fmt.Errorf("chart style %d is outside the supported range 1..48: %w", style, toolkiterrors.ErrInvalidChartStyle)
	}
	return nil
}

// ValidateChartType validates a chart type name. The name must be one of the
// engine's 81 supported chart types.
//
// Parameters:
//   - chartType: The chart type to validate.
//
// Returns:
//   - error: An error if the chart type is not recognized.
//
// Example:
//
//	if err := editor.ValidateChartType(editor.ChartTypeColumn); err != nil {
//	    log.Printf("Invalid chart type: %v", err)
//	}
func ValidateChartType(chartType ChartType) error {
	_, err := cells.ResolveChartType(string(chartType))
	if err != nil {
		return fmt.Errorf("chart type %q is not recognized: %w (hint: use one of the ChartType constants or any of the engine's 81 chart type names)", chartType, err)
	}
	return nil
}

// ValidateLegendPosition validates a legend position name.
//
// Parameters:
//   - position: The legend position to validate.
//
// Returns:
//   - error: An error if the position is not recognized.
//
// Example:
//
//	if err := editor.ValidateLegendPosition(editor.ChartLegendBottom); err != nil {
//	    log.Printf("Invalid legend position: %v", err)
//	}
func ValidateLegendPosition(position ChartLegendPosition) error {
	_, err := cells.ResolveLegendPosition(string(position))
	if err != nil {
		return fmt.Errorf("legend position %q is not recognized: %w (hint: use one of the ChartLegend constants)", position, err)
	}
	return nil
}

// SuggestDataRange provides guidance on choosing a data range for a chart based
// on the data layout. It returns a human-readable suggestion.
//
// Parameters:
//   - dataRows: The number of rows of data (excluding headers).
//   - dataCols: The number of columns of data (excluding category labels).
//   - hasCategories: Whether the first column contains category labels.
//
// Returns:
//   - string: A suggestion for the data range in A1 notation.
//
// Example:
//
//	suggestion := editor.SuggestDataRange(10, 3, true)
//	// Returns "A1:D11" (10 data rows + 1 header, 3 data cols + 1 category col)
func SuggestDataRange(dataRows, dataCols int, hasCategories bool) string {
	if dataRows <= 0 || dataCols <= 0 {
		return ""
	}

	totalRows := dataRows + 1 // Include header row
	totalCols := dataCols
	if hasCategories {
		totalCols++ // Include category column
	}

	// Convert to Excel column letters
	startCol := 'A'
	endCol := rune('A' + totalCols - 1)

	return fmt.Sprintf("%c1:%c%d", startCol, endCol, totalRows)
}

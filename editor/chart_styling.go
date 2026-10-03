package editor

import (
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// WithChartTitleFont creates a ChartAction that sets the chart title's font
// properties. This allows fine-grained control over the title's appearance.
//
// Parameters:
//   - fontName: The font name (e.g., "Arial", "Times New Roman"). Empty means no change.
//   - fontSize: The font size in points. Zero means no change.
//   - bold: Whether the font is bold. Only applied if fontName or fontSize is set.
//
// Returns:
//   - ChartAction: A function that sets the title font.
//
// Example:
//
//	editor.InChart(0, editor.WithChartTitleFont("Arial", 14, true))
func WithChartTitleFont(fontName string, fontSize int, bold bool) ChartAction {
	return func(chart *asposecells.Chart) error {
		title, err := cells.ChartTitle(chart)
		if err != nil {
			return err
		}
		font, err := engine.Derive(title.GetFont())
		if err != nil {
			return err
		}
		if fontName != "" {
			if err := font.SetName_String(fontName); err != nil {
				return err
			}
		}
		if fontSize > 0 {
			if err := font.SetSize(int32(fontSize)); err != nil {
				return err
			}
		}
		if fontName != "" || fontSize > 0 {
			if err := font.SetIsBold(bold); err != nil {
				return err
			}
		}
		return nil
	}
}

// WithChartTitleColor creates a ChartAction that sets the chart title's font
// color.
//
// Parameters:
//   - color: The color as a hex string (e.g., "#FF0000" for red), a color
//     name, a Go color.Color, or an ARGB int. See WithFontColor.
//
// Returns:
//   - ChartAction: A function that sets the title color.
//
// Example:
//
//	editor.InChart(0, editor.WithChartTitleColor("#FF0000"))
func WithChartTitleColor(color interface{}) ChartAction {
	return func(chart *asposecells.Chart) error {
		title, err := cells.ChartTitle(chart)
		if err != nil {
			return err
		}
		font, err := engine.Derive(title.GetFont())
		if err != nil {
			return err
		}
		v, err := resolveColor(color)
		if err != nil {
			return err
		}
		return font.SetColor(v)
	}
}

// WithChartLegendFont creates a ChartAction that sets the chart legend's font
// properties.
//
// Parameters:
//   - fontName: The font name. Empty means no change.
//   - fontSize: The font size in points. Zero means no change.
//   - bold: Whether the font is bold.
//
// Returns:
//   - ChartAction: A function that sets the legend font.
//
// Example:
//
//	editor.InChart(0, editor.WithChartLegendFont("Calibri", 10, false))
func WithChartLegendFont(fontName string, fontSize int, bold bool) ChartAction {
	return func(chart *asposecells.Chart) error {
		legend, err := cells.ChartLegend(chart)
		if err != nil {
			return err
		}
		font, err := engine.Derive(legend.GetFont())
		if err != nil {
			return err
		}
		if fontName != "" {
			if err := font.SetName_String(fontName); err != nil {
				return err
			}
		}
		if fontSize > 0 {
			if err := font.SetSize(int32(fontSize)); err != nil {
				return err
			}
		}
		if fontName != "" || fontSize > 0 {
			if err := font.SetIsBold(bold); err != nil {
				return err
			}
		}
		return nil
	}
}

// WithChartSeriesColor creates a ChartAction that sets the color of a specific
// chart series.
//
// Parameters:
//   - seriesIndex: The zero-based index of the series.
//   - color: The color as a hex string, a color name, a Go color.Color, or
//     an ARGB int. See WithFontColor.
//
// Returns:
//   - ChartAction: A function that sets the series color.
//
// Example:
//
//	editor.InChart(0, editor.WithChartSeriesColor(0, "#00FF00"))
func WithChartSeriesColor(seriesIndex int, color interface{}) ChartAction {
	return func(chart *asposecells.Chart) error {
		series, err := cells.ChartSeries(chart)
		if err != nil {
			return err
		}
		s, err := series.Get(int32(seriesIndex))
		if err != nil {
			return err
		}
		area, err := s.GetArea()
		if err != nil {
			return err
		}
		v, err := resolveColor(color)
		if err != nil {
			return err
		}
		return area.SetForegroundColor(v)
	}
}

// WithChartSeriesName creates a ChartAction that sets the name of a specific
// chart series. The name appears in the legend.
//
// Parameters:
//   - seriesIndex: The zero-based index of the series.
//   - name: The series name.
//
// Returns:
//   - ChartAction: A function that sets the series name.
//
// Example:
//
//	editor.InChart(0, editor.WithChartSeriesName(0, "Q1 Sales"))
func WithChartSeriesName(seriesIndex int, name string) ChartAction {
	return func(chart *asposecells.Chart) error {
		series, err := cells.ChartSeries(chart)
		if err != nil {
			return err
		}
		s, err := series.Get(int32(seriesIndex))
		if err != nil {
			return err
		}
		return s.SetName(name)
	}
}

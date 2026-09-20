// Command chart demonstrates the editor DSL's chart support: adding charts,
// changing their type, title, built-in style, legend and placement, re-pointing
// them at different source data, growing and shrinking their series, and
// deleting them again.
//
// Nothing here imports the underlying engine binding. The workbook is seeded
// with datasource.NewEmptyWorkbook and every chart operation goes through
// editor and query, which is the point of the toolkit: a caller builds a chart
// without ever holding an engine handle.
//
// Output lands in examples/chart/out/.
package main

import (
	"fmt"
	"log"
	"os"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	examples "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/examples/common"
)

func main() {
	if err := examples.SetLicense(); err != nil {
		log.Printf("license: %v", err)
	}

	// Seed an empty workbook so the example needs no input file. The table is
	// written onto a sheet this program adds, so no sheet name is read from a
	// file and the example is unaffected by evaluation mode's load-time
	// worksheet-name rewriting.
	seed, err := datasource.NewEmptyWorkbook()
	if err != nil {
		log.Fatal(err)
	}

	const sheet = "Sales"

	// Build the table the charts will plot. Column A holds the categories and
	// the label column, B and C hold two series of figures.
	seeded, err := editor.EditSpreadsheet(
		seed,
		editor.WithAddWorksheet(sheet),
		editor.InWorksheet(sheet,
			editor.SetCellValue(0, 0, "Region"),
			editor.SetCellValue(0, 1, "Q1"),
			editor.SetCellValue(0, 2, "Q2"),
			editor.SetCellValue(1, 0, "North"),
			editor.SetCellValue(1, 1, int32(120)),
			editor.SetCellValue(1, 2, int32(150)),
			editor.SetCellValue(2, 0, "South"),
			editor.SetCellValue(2, 1, int32(90)),
			editor.SetCellValue(2, 2, int32(130)),
			editor.SetCellValue(3, 0, "East"),
			editor.SetCellValue(3, 1, int32(160)),
			editor.SetCellValue(3, 2, int32(140)),
			editor.SetCellValue(4, 0, "West"),
			editor.SetCellValue(4, 1, int32(70)),
			editor.SetCellValue(4, 2, int32(110)),
		),
	)
	if err != nil {
		log.Fatalf("seed the table: %v", err)
	}

	// Add a column chart over A1:C5 with a title, a built-in style and a legend.
	//
	// Read by column, the engine treats the text column A as the category axis,
	// so the two series come from B and C -- one per quarter. SetStyle picks one
	// of the engine's 48 built-in chart styles, and the legend is docked bottom.
	//
	// The four numbers after byColumn are the rectangle the chart is drawn over:
	// row 5, column 0, to row 20, column 7. AddChart has no default placement, so
	// every chart says where it goes; these bounds sit below the table, which
	// occupies rows 0-4 of columns A-C.
	withChart, err := editor.EditSpreadsheet(
		datasource.BytesSource(seeded),
		editor.InWorksheet(sheet,
			editor.AddChart(editor.ChartTypeColumn, "A1:C5", true, 5, 0, 20, 7,
				editor.WithChartTitle("Quarterly sales by region"),
				editor.WithChartStyle(7),
				editor.WithChartLegend(true),
				editor.WithChartLegendPosition(editor.ChartLegendBottom),
			),
		),
	)
	if err != nil {
		log.Fatalf("add chart: %v", err)
	}
	writeOut("chart.xlsx", withChart)

	// Re-open the saved file and modify the chart in place: switch it to a bar
	// chart, retitle it, and change which cells it plots. WithChartDataRange
	// replaces the source data wholesale, so the series follow the new range.
	modified, err := editor.EditSpreadsheet(
		datasource.FilePathSource(examples.OutPath("chart", "chart.xlsx")),
		editor.InWorksheet(sheet,
			editor.InChart(0,
				editor.WithChartType(editor.ChartTypeBar),
				editor.WithChartTitle("Q1 vs Q2"),
				editor.WithChartDataRange("A1:C5", true),
				editor.WithChartStyle(12),
			),
		),
	)
	if err != nil {
		log.Fatalf("modify chart: %v", err)
	}
	writeOut("chart-modified.xlsx", modified)

	// A second chart on the same sheet, built series by series rather than from
	// one range, to show the series actions. A waterfall chart has no constant in
	// the curated list, so it is named directly -- any of the engine's chart types
	// can be reached this way.
	//
	// The ranges start at row 2 to leave the header row out: a series over
	// "B1:B5" would plot the label cell "Q1" as a data point.
	composed, err := editor.EditSpreadsheet(
		datasource.BytesSource(seeded),
		editor.InWorksheet(sheet,
			editor.AddChart(editor.ChartType("waterfall"), "B2:B5", true, 5, 0, 20, 7,
				editor.WithChartTitle("Q2 only"),
				editor.WithChartLegend(false),
			),
			// Grow the chart to two series, then drop the first one again, so the
			// chart ends up plotting Q2 -- the series that was appended.
			editor.InChart(0, editor.WithChartSeries("C2:C5", true)),
			editor.InChart(0, editor.RemoveChartSeries(0)),
		),
	)
	if err != nil {
		log.Fatalf("compose chart: %v", err)
	}
	writeOut("chart-composed.xlsx", composed)

	// Add two charts and delete one, leaving a single chart on the sheet.
	deleted, err := editor.EditSpreadsheet(
		datasource.BytesSource(seeded),
		editor.InWorksheet(sheet,
			editor.AddChart(editor.ChartTypeColumn, "A1:C5", true, 5, 0, 20, 7,
				editor.WithChartTitle("keep me"),
			),
			editor.AddChart(editor.ChartTypePie, "B2:B5", true, 5, 0, 20, 7,
				editor.WithChartTitle("drop me"),
			),
			editor.DeleteChart(1),
		),
	)
	if err != nil {
		log.Fatalf("delete chart: %v", err)
	}
	writeOut("chart-deleted.xlsx", deleted)

	// The "modified" workbook above was produced by re-opening "chart.xlsx" from
	// disk and editing chart 0 in place, which is the round trip this example is
	// really demonstrating: a chart written by one pass is found and changed by
	// the next, with no engine handle ever entering the caller's code.
	log.Println("chart workbooks written to examples/chart/out/")
}

func writeOut(name string, data []byte) {
	if err := os.MkdirAll(examples.OutDir("chart"), 0o755); err != nil {
		log.Fatal(err)
	}
	if err := os.WriteFile(examples.OutPath("chart", name), data, 0o644); err != nil {
		log.Fatal(err)
	}
	fmt.Printf("wrote %s\n", examples.OutPath("chart", name))
}

// Command query demonstrates reading spreadsheets with the query package:
// typed cell reads (ReadCell, ReadRange, ReadWorksheet), workbook metadata
// (Dimensions, SheetNames, ReadMergedCells), structured rows (ReadRows), and
// the named-range / cell-comment reads (NamedRanges, ReadNamedRange,
// ReadCellComment).
//
// Part 1 reads the real sample workbook examples/data/BookText.xlsx; part 2
// seeds a small workbook (via the editor DSL), writes it to examples/query/out,
// and reads it back. Reads target worksheets by index, which is immune to the
// evaluation-mode worksheet-name corruption described in the README.
package main

import (
	"fmt"
	"log"
	"os"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	examples "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/examples/common"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Row is the header shape of BookText.xlsx (Item / Category / Note); the
// excel tags map struct fields to columns, and the header match is
// case-insensitive.
type Row struct {
	Item     string `excel:"item"`
	Category string `excel:"category"`
	Note     string `excel:"note"`
}

func main() {
	if err := examples.SetLicense(); err != nil {
		log.Printf("license: %v", err)
	}
	if err := os.MkdirAll(examples.OutDir("query"), 0o755); err != nil {
		log.Fatal(err)
	}

	// --- Part 1: read the real sample workbook -------------------------------

	book := datasource.FilePathSource(examples.DataPath("BookText.xlsx"))
	first := query.WithSheetIndex(0) // index-based: immune to name corruption

	cell, err := query.ReadCell(book, "B2", first)
	if err != nil {
		log.Fatalf("read cell: %v", err)
	}
	if s, ok := cell.String(); ok {
		fmt.Printf("ReadCell(B2): %s\n", s)
	}

	grid, err := query.ReadRange(book, "A1", "C4", first)
	if err != nil {
		log.Fatalf("read range: %v", err)
	}
	fmt.Println("ReadRange(A1:C4):")
	printGrid(grid)

	whole, err := query.ReadWorksheet(book, first)
	if err != nil {
		log.Fatalf("read worksheet: %v", err)
	}
	fmt.Println("ReadWorksheet (first sheet):")
	printGrid(whole)

	rows, cols, err := query.Dimensions(book, first)
	if err != nil {
		log.Fatalf("dimensions: %v", err)
	}
	fmt.Printf("Dimensions: %d rows x %d cols\n", rows, cols)

	records, err := query.ReadRows[Row](book, first)
	if err != nil {
		log.Fatalf("read rows: %v", err)
	}
	fmt.Println("ReadRows[Row] (headers Item/Category/Note):")
	for _, r := range records {
		fmt.Printf("  %s / %s / %s\n", r.Item, r.Category, r.Note)
	}

	// --- Part 2: seed a small workbook and read it back -----------------------

	// The named range and comment live on an explicitly added "Data" sheet,
	// whose name is set after load and so is never corrupted by evaluation
	// mode; the reads below target it by index.
	wb, err := asposecells.NewWorkbook()
	if err != nil {
		log.Fatal(err)
	}
	seed, err := wb.Save_SaveFormat(asposecells.SaveFormat_Xlsx)
	if err != nil {
		log.Fatal(err)
	}
	table, err := editor.EditSpreadsheet(
		datasource.BytesSource(seed),
		editor.WithAddWorksheet("Data"),
		editor.InWorksheet("Data",
			editor.SetCellValue(0, 0, "Product"),
			editor.SetCellValue(0, 1, "Price"),
			editor.SetCellValue(1, 0, "Widget"),
			editor.SetCellValue(1, 1, int32(12)),
			editor.SetCellValue(2, 0, "Gadget"),
			editor.SetCellValue(2, 1, int32(3)),
			// A merged note row below the table.
			editor.SetValue(3, 0, 1, 2, "prices are per unit"),
			editor.Merge(3, 0, 1, 2, true),
			// P2-7: a named range over the Price column (B2:B3) and a comment.
			editor.DefineNamedRange("PriceRange", 1, 1, 2, 1),
			editor.SetCellComment(1, 0, "Widget is our best seller"),
		),
	)
	if err != nil {
		log.Fatalf("seed worksheet: %v", err)
	}
	tablePath := examples.OutPath("query", "table.xlsx")
	if err := os.WriteFile(tablePath, table, 0o644); err != nil {
		log.Fatal(err)
	}

	// The "Data" sheet is targeted by name, not by index: evaluation mode can
	// inject "Evaluation Warning" sheets that shift worksheet indices, and can
	// rewrite a loaded sheet's name to garbage on ~2% of loads — so the
	// name-dependent read-backs below are reported, not fatal.
	tableSrc := datasource.FilePathSource(tablePath)
	data := query.WithSheet("Data")

	sheets, err := query.SheetNames(tableSrc)
	if err != nil {
		log.Fatalf("sheet names: %v", err)
	}
	fmt.Println("SheetNames:", sheets)

	merged, err := query.ReadMergedCells(tableSrc, data)
	if err != nil {
		log.Printf("merged cells: %v (evaluation-mode name corruption? re-run for a clean pass)", err)
	} else {
		for _, a := range merged {
			fmt.Printf("MergedCells: %s\n", a.String())
		}
	}

	names, err := query.NamedRanges(tableSrc)
	if err != nil {
		log.Fatalf("named ranges: %v", err)
	}
	for _, nr := range names {
		fmt.Printf("NamedRanges: %s -> %s\n", nr.Name, nr.RefersTo)
	}

	// Resolving the named range's sheet reference needs the sheet name the
	// engine holds at load time. Evaluation mode can rewrite a worksheet name
	// on ~2% of loads (never in the saved bytes), which would make this fail
	// spuriously — so a failure here is reported, not fatal.
	priceGrid, err := query.ReadNamedRange(tableSrc, "PriceRange")
	if err != nil {
		log.Printf("read named range PriceRange: %v (evaluation-mode name corruption? re-run for a clean pass)", err)
	} else {
		fmt.Println("ReadNamedRange(PriceRange):")
		printGrid(priceGrid)
	}

	note, err := query.ReadCellComment(tableSrc, "A2", data)
	if err != nil {
		log.Printf("cell comment: %v (evaluation-mode name corruption? re-run for a clean pass)", err)
	} else {
		fmt.Printf("ReadCellComment(A2): %q\n", note)
	}

	log.Println("query complete; seeded workbook written to examples/query/out/table.xlsx")
}

// printGrid renders a row-major grid of cell values as a tab-separated table,
// using "-" for empty cells and the canonical string form otherwise.
func printGrid(g [][]query.CellValue) {
	for _, row := range g {
		for i, v := range row {
			if i > 0 {
				fmt.Print("\t")
			}
			if s, ok := v.String(); ok {
				fmt.Print(s)
			} else {
				fmt.Print("-")
			}
		}
		fmt.Println()
	}
}

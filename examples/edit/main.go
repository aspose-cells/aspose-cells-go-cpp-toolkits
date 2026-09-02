// Command edit demonstrates programmatic spreadsheet editing with the editor
// DSL: writing values, styling ranges, the full worksheet operation set
// (insert, delete, merge, unmerge, clear, and range edits), and the named
// range / cell comment / encryption actions. The edited workbook is read back
// with the query package to show the metadata round-trips, and an encrypted
// copy is written alongside it.
//
// The example seeds an empty workbook, edits it, and writes the results to
// examples/edit/out/.
package main

import (
	"fmt"
	"log"
	"os"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/core"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	examples "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/examples/common"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

func main() {
	if p := os.Getenv("LicensePath"); p != "" {
		if err := core.SetLicense(p); err != nil {
			log.Printf("license: %v", err)
		}
	}

	// Seed an empty workbook so the example runs without a data file.
	wb, err := asposecells.NewWorkbook()
	if err != nil {
		log.Fatal(err)
	}
	seed, err := wb.Save_SaveFormat(asposecells.SaveFormat_Xlsx)
	if err != nil {
		log.Fatal(err)
	}

	// The table, named range, and comment live on a "Data" sheet added after
	// the workbook is loaded. Evaluation mode can rewrite a worksheet's name to
	// garbage on ~2% of loads (never in the saved bytes), and the named range's
	// stored reference embeds the sheet name — so the range is defined on a
	// sheet whose name is set by this call, not one read from the file.
	edited, err := editor.EditSpreadsheet(
		datasource.BytesSource(seed),
		editor.WithAddWorksheet("Data"),
		editor.WithActiveSheet("Data"),
		editor.InWorksheet("Data",
			// Build a small table on the sheet.
			editor.SetCellValue(0, 0, "Item"),
			editor.SetCellValue(0, 1, "Qty"),
			editor.SetCellValue(0, 2, "In stock"),
			editor.SetCellValue(1, 0, "Widget"),
			editor.SetCellValue(1, 1, int32(12)),
			editor.SetCellValue(1, 2, true),
			editor.SetCellValue(2, 0, "Gadget"),
			editor.SetCellValue(2, 1, int32(3)),
			editor.SetCellValue(2, 2, false),
			// Bold, 14pt red header across the first row.
			editor.SetStyle(0, 0, 1, 3,
				editor.WithFontName("Arial"),
				editor.WithFontSize(14),
				editor.WithFontIsBold(true),
				editor.WithFontColor("red"),
			),
			// Structure edits on the rows below the table.
			editor.InsertRows(4, 1, true),
			editor.InsertColumns(3, 1, true),
			editor.SetValue(5, 0, 1, 2, "Merged"),
			editor.Merge(5, 0, 1, 2, true),
			editor.UnMerge(5, 0, 1, 2),
			editor.SetValue(6, 0, 1, 4, "Four columns wide"),
			editor.ClearContents(6, 0, 1, 4),
			editor.ClearFormats(6, 0, 1, 4),
			editor.DeleteRows(6, 1, true),
			editor.DeleteRange(4, 4, 1, 1, "up"),
			editor.DeleteBlankRows(),
			editor.DeleteBlankColumns(),
			// A named range over the Qty column (B2:B3) and a comment on A2.
			editor.DefineNamedRange("QtyRange", 1, 1, 2, 1),
			editor.SetCellComment(1, 0, "Widget is our best seller"),
		),
		// Add a second sheet and write to it by name.
		editor.WithAddWorksheet("Summary"),
		editor.InWorksheet("Summary", editor.SetCellValue(0, 0, "Total")),
		// Default workbook font.
		editor.InDefaultStyle(editor.WithFontName("Calibri")),
	)
	if err != nil {
		log.Fatalf("edit spreadsheet: %v", err)
	}
	if err := os.MkdirAll(examples.OutDir("edit"), 0o755); err != nil {
		log.Fatal(err)
	}
	editedPath := examples.OutPath("edit", "edited.xlsx")
	if err := os.WriteFile(editedPath, edited, 0o644); err != nil {
		log.Fatal(err)
	}

	// Read the file back and push a further edit into a copy, then write it out.
	editedCopy, err := editor.EditSpreadsheet(
		datasource.FilePathSource(editedPath),
		editor.InWorksheet(1, editor.SetCellValue(0, 0, "Touched")), // 1 = "Data"
	)
	if err != nil {
		log.Fatalf("edit copy: %v", err)
	}
	if err := os.WriteFile(examples.OutPath("edit", "edited-copy.xlsx"), editedCopy, 0o644); err != nil {
		log.Fatal(err)
	}

	// Read the saved workbook back with the query package to show the named
	// range, its cells, and the comment survived the round-trip.
	editedSrc := datasource.FilePathSource(editedPath)
	names, err := query.NamedRanges(editedSrc)
	if err != nil {
		log.Fatalf("read named ranges: %v", err)
	}
	for _, nr := range names {
		fmt.Printf("named range %s: %s -> %s\n", nr.Name, nr.RefersTo, nr.Area.String())
	}
	// The comment lives on the "Data" sheet, targeted by name because evaluation
	// mode can inject "Evaluation Warning" sheets that shift worksheet indices.
	// Resolving it also depends on the sheet name the engine holds at load time,
	// which evaluation mode can rewrite to garbage on ~2% of loads — so the
	// name-dependent read-backs below are reported, not fatal.
	note, err := query.ReadCellComment(editedSrc, "A2", query.WithSheet("Data"))
	if err != nil {
		log.Printf("read cell comment: %v (evaluation-mode name corruption? re-run for a clean pass)", err)
	} else {
		fmt.Printf("comment on A2: %q\n", note)
	}
	qty, err := query.ReadNamedRange(editedSrc, "QtyRange")
	if err != nil {
		log.Printf("read named range QtyRange: %v (evaluation-mode name corruption? re-run for a clean pass)", err)
	} else {
		fmt.Print("QtyRange:")
		for _, row := range qty {
			for _, cell := range row {
				if v, ok := cell.Int(); ok {
					fmt.Printf(" %d", v)
				}
			}
		}
		fmt.Println()
	}

	// Encrypt the workbook so the saved file requires a password to open.
	// The toolkit loader does not expose a load-time password yet, so an
	// encrypted file is only ever written here, not read back.
	encrypted, err := editor.EditSpreadsheet(
		datasource.BytesSource(edited),
		editor.Encrypt("hunter2"),
	)
	if err != nil {
		log.Fatalf("encrypt: %v", err)
	}
	if err := os.WriteFile(examples.OutPath("edit", "edited-encrypted.xlsx"), encrypted, 0o644); err != nil {
		log.Fatal(err)
	}

	log.Println("edited workbooks written to examples/edit/out/edited.xlsx, edited-copy.xlsx, and edited-encrypted.xlsx")
}

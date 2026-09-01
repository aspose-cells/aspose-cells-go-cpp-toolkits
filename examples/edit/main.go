// Command edit demonstrates programmatic spreadsheet editing with the editor
// DSL: writing values, styling ranges, and the full worksheet operation set
// (insert, delete, merge, unmerge, clear, and range edits).
//
// The example seeds an empty workbook, edits it, and writes the result to
// examples/edit/out/edited.xlsx.
package main

import (
	"log"
	"os"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/core"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	examples "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/examples/common"
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

	edited, err := editor.EditSpreadsheet(
		datasource.BytesSource(seed),
		// Activate a sheet and build a small table on it.
		editor.WithActiveSheet("Sheet1"),
		editor.InWorksheet(0,
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
		editor.InWorksheet(0, editor.SetCellValue(0, 0, "Touched")),
	)
	if err != nil {
		log.Fatalf("edit copy: %v", err)
	}
	if err := os.WriteFile(examples.OutPath("edit", "edited-copy.xlsx"), editedCopy, 0o644); err != nil {
		log.Fatal(err)
	}
	log.Println("edited workbooks written to examples/edit/out/edited.xlsx and examples/edit/out/edited-copy.xlsx")
}

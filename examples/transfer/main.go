// Command transfer demonstrates exporting spreadsheet data to structured
// formats (XML, JSON) and importing CSV data back into a spreadsheet, using
// every transfer entry point:
//
//   - ExportWorksheetToJsonFile:   sheet -> JSON file
//   - ExportRangeToJsonFile:       cell range -> JSON file
//   - ExportSpreadsheetToXml:      workbook -> XML []byte
//   - ExportSpreadsheetToXmlFile:  workbook -> XML file
//   - ImportCSVFile:               CSV -> worksheet
//   - ImportJsonFile:              JSON -> worksheet
//   - ImportXMLFile:               XML -> worksheet
//
// It seeds an in-memory workbook; the import samples come from examples/data
// and all outputs go to examples/transfer/out.
package main

import (
	"log"
	"os"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/core"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	examples "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/examples/common"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/transfer"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

func main() {
	if p := os.Getenv("LicensePath"); p != "" {
		core.SetLicense(p)
	}
	if err := os.MkdirAll(examples.OutDir("transfer"), 0o755); err != nil {
		log.Fatal(err)
	}

	// Seed an in-memory workbook with a small table. The default first sheet's
	// name is unreliable in evaluation mode (the engine can corrupt it), so the
	// table goes on an explicitly named "Data" sheet; the imports target a
	// worksheet named "Imported", created up front.
	wb, err := asposecells.NewWorkbook()
	if err != nil {
		log.Fatal(err)
	}
	empty, err := wb.Save_SaveFormat(asposecells.SaveFormat_Xlsx)
	if err != nil {
		log.Fatal(err)
	}
	seeded, err := editor.EditSpreadsheet(
		datasource.BytesSource(empty),
		editor.WithAddWorksheet("Data"),
		editor.InWorksheet("Data",
			editor.SetCellValue(0, 0, "ID"),
			editor.SetCellValue(0, 1, "Name"),
			editor.SetCellValue(1, 0, int32(1)),
			editor.SetCellValue(1, 1, "Alpha"),
			editor.SetCellValue(2, 0, int32(2)),
			editor.SetCellValue(2, 1, "Beta"),
		),
		editor.WithAddWorksheet("Imported"),
	)
	if err != nil {
		log.Fatal(err)
	}
	source := datasource.BytesSource(seeded)

	// Persist the seed so the file-based entry points have a real input.
	seedPath := examples.OutPath("transfer", "seed.xlsx")
	if err := os.WriteFile(seedPath, seeded, 0o644); err != nil {
		log.Fatal(err)
	}

	// Export the "Data" sheet as JSON straight from file to file.
	if err := transfer.ExportWorksheetToJsonFile(seedPath, "Data", examples.OutPath("transfer", "sheet.json")); err != nil {
		log.Fatalf("export worksheet json: %v", err)
	}

	// Export a cell range as JSON.
	if err := transfer.ExportRangeToJsonFile(seedPath, "Data", "A1", "B3", examples.OutPath("transfer", "range.json")); err != nil {
		log.Fatalf("export range json: %v", err)
	}

	// Export the whole workbook as XML, both to bytes and to a file.
	xmlData, err := transfer.ExportSpreadsheetToXml(source, "InventoryMap")
	if err != nil {
		log.Fatalf("export xml: %v", err)
	}
	if err := os.WriteFile(examples.OutPath("transfer", "inventory.xml"), xmlData, 0o644); err != nil {
		log.Fatal(err)
	}
	if err := transfer.ExportSpreadsheetToXmlFile(seedPath, "InventoryMap", examples.OutPath("transfer", "inventory-file.xml")); err != nil {
		log.Fatalf("export xml file: %v", err)
	}

	// Import CSV, JSON, and XML data into fresh copies of the workbook. The
	// sample data files live in examples/data.
	if err := transfer.ImportCSVFile(seedPath, examples.DataPath("BookCsvDuplicateData.csv"), "Imported", 0, 0, true, ",", examples.OutPath("transfer", "imported-csv.xlsx")); err != nil {
		log.Fatalf("import csv: %v", err)
	}
	if err := transfer.ImportJsonFile(seedPath, examples.DataPath("importdata.json"), "Imported", 0, 0, examples.OutPath("transfer", "imported-json.xlsx")); err != nil {
		log.Fatalf("import json: %v", err)
	}
	if err := transfer.ImportXMLFile(seedPath, examples.DataPath("data.xml"), "Imported", 0, 0, examples.OutPath("transfer", "imported-xml.xlsx")); err != nil {
		log.Fatalf("import xml: %v", err)
	}

	log.Println("transfer complete; outputs written to examples/transfer/out")
}

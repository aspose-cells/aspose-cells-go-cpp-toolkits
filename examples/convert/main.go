// Command convert demonstrates converting a spreadsheet between formats using
// the converter package and the saveoptions format packages.
//
// It converts the sample workbook examples/data/BookText.xlsx through the three
// conversion entry points:
//
//   - ConvertSpreadsheet:       spreadsheet -> []byte (PDF here, with options)
//   - ConvertToWriter:          spreadsheet -> io.Writer (Markdown)
//   - ConvertSpreadsheetToFile: file -> file, format from the extension (CSV/JSON)
//
// Outputs are written to examples/convert/out.
package main

import (
	"log"
	"os"
	"path/filepath"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/core"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	examples "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/examples/common"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	_ "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/register"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/markdown"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pdf"
)

func main() {
	if p := os.Getenv("LicensePath"); p != "" {
		core.SetLicense(p)
	}
	if err := os.MkdirAll(examples.OutDir("convert"), 0o755); err != nil {
		log.Fatal(err)
	}

	// The sample workbook lives in examples/data so the command is
	// self-contained; pass a path as an argument to convert your own file.
	input := examples.DataPath("BookText.xlsx")
	if len(os.Args) > 1 {
		input = os.Args[1]
	}
	source := datasource.FilePathSource(input)

	// Spreadsheet -> []byte: PDF with a functional option.
	pdfData, err := converter.ConvertSpreadsheet(source, pdf.New(pdf.WithOnePagePerSheet(true)))
	if err != nil {
		log.Fatalf("convert to pdf: %v", err)
	}
	if err := os.WriteFile(examples.OutPath("convert", "report.pdf"), pdfData, 0o644); err != nil {
		log.Fatal(err)
	}

	// Spreadsheet -> io.Writer: Markdown streamed straight to a file.
	file, err := os.Create(examples.OutPath("convert", "report.md"))
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()
	if err := converter.ConvertToWriter(source, file, markdown.New(markdown.WithClearData(true))); err != nil {
		log.Fatalf("convert to markdown: %v", err)
	}

	// File -> file: the output format is inferred from the extension.
	for _, out := range []string{examples.OutPath("convert", "report.csv"), examples.OutPath("convert", "report.json")} {
		if err := converter.ConvertSpreadsheetToFile(input, out); err != nil {
			log.Fatalf("convert to %s: %v", filepath.Base(out), err)
		}
	}

	// Resolve the format from the extension, then write the bytes yourself.
	ext := filepath.Ext(input)[1:]
	opt := formats.Get(ext)
	if opt == nil {
		log.Fatalf("no registered format for %q", ext)
	}
	copyData, err := converter.ConvertSpreadsheet(source, opt)
	if err != nil {
		log.Fatalf("convert to %s: %v", ext, err)
	}
	if err := os.WriteFile(examples.OutPath("convert", "copy."+ext), copyData, 0o644); err != nil {
		log.Fatal(err)
	}

	log.Println("conversion complete; outputs written to examples/convert/out")
}

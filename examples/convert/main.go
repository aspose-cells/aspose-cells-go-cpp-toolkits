// Command convert demonstrates converting a spreadsheet between formats using
// the converter package and the saveoptions format packages.
//
// It converts the sample workbook examples/data/BookText.xlsx through the
// sink-based converter.Convert, choosing the output shape by picking a sink:
//
//   - datasource.BytesSink:   spreadsheet -> []byte (PDF with options)
//   - datasource.WriterSink:  spreadsheet -> io.Writer (Markdown)
//   - datasource.FilePathSink:file -> file, format from the extension (CSV/JSON)
//
// and through the three earlier converter entry points (kept as deprecated thin
// wrappers) so every converter API is exercised:
//
//   - ConvertSpreadsheet:       spreadsheet -> []byte
//   - ConvertToWriter:          spreadsheet -> io.Writer
//   - ConvertSpreadsheetToFile: file -> file, format from the extension
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
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/ooxml"
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

	// Spreadsheet -> []byte: PDF with a functional option via a BytesSink.
	var pdfSink datasource.BytesSink
	if err := converter.Convert(source, pdf.New(pdf.WithOnePagePerSheet(true)), &pdfSink); err != nil {
		log.Fatalf("convert to pdf: %v", err)
	}
	if err := os.WriteFile(examples.OutPath("convert", "report.pdf"), pdfSink.Bytes(), 0o644); err != nil {
		log.Fatal(err)
	}

	// Spreadsheet -> io.Writer: Markdown streamed straight to a file via a
	// WriterSink.
	file, err := os.Create(examples.OutPath("convert", "report.md"))
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()
	if err := converter.Convert(source, markdown.New(markdown.WithClearData(true)), datasource.NewWriterSink(file)); err != nil {
		log.Fatalf("convert to markdown: %v", err)
	}

	// File -> file: resolve the output format from the extension, then hand a
	// FilePathSink to Convert.
	for _, out := range []string{examples.OutPath("convert", "report.csv"), examples.OutPath("convert", "report.json")} {
		ext := filepath.Ext(out)[1:]
		opt := formats.Get(ext)
		if opt == nil {
			log.Fatalf("no registered format for %q", ext)
		}
		if err := converter.Convert(source, opt, datasource.FilePathSink(out)); err != nil {
			log.Fatalf("convert to %s: %v", ext, err)
		}
	}

	// Deprecated wrappers, kept for backward compatibility. New code should use
	// converter.Convert with an appropriate sink as above.

	// ConvertSpreadsheet returns the converted content as []byte (XLSX here,
	// preserving the original workbook structure).
	bytesData, err := converter.ConvertSpreadsheet(source, ooxml.New())
	if err != nil {
		log.Fatalf("convert spreadsheet: %v", err)
	}
	if err := os.WriteFile(examples.OutPath("convert", "copy.xlsx"), bytesData, 0o644); err != nil {
		log.Fatal(err)
	}

	// ConvertToWriter writes straight into an io.Writer (CSV here).
	csvFile, err := os.Create(examples.OutPath("convert", "legacy.csv"))
	if err != nil {
		log.Fatal(err)
	}
	csvOpt := formats.Get("csv")
	if csvOpt == nil {
		log.Fatal(`no registered format for "csv"`)
	}
	if err := converter.ConvertToWriter(source, csvFile, csvOpt); err != nil {
		log.Fatalf("convert to writer: %v", err)
	}
	csvFile.Close()

	// ConvertSpreadsheetToFile infers the output format from the extension.
	if err := converter.ConvertSpreadsheetToFile(input, examples.OutPath("convert", "legacy.html")); err != nil {
		log.Fatalf("convert to file: %v", err)
	}

	log.Println("conversion complete; outputs written to examples/convert/out")
}

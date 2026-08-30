// Command merge-split demonstrates combining several workbooks into one and
// splitting a workbook back into per-worksheet files, using the two manipulator
// entry points with sink-chosen output shapes:
//
//   - manipulator.Merge:  []DataSource -> BytesSink / WriterSink / FilePathSink
//   - manipulator.Split:  DataSource -> ZipSink (archive) / FolderSink (files)
//
// It seeds two workbooks in memory; all outputs go to examples/merge-split/out.
package main

import (
	"archive/zip"
	"bytes"
	"log"
	"os"
	"path/filepath"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/core"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	examples "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/examples/common"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/manipulator"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/html"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/ooxml"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

func main() {
	if p := os.Getenv("LicensePath"); p != "" {
		core.SetLicense(p)
	}
	out := examples.OutDir("merge-split")
	if err := os.MkdirAll(out, 0o755); err != nil {
		log.Fatal(err)
	}
	sheetsDir := filepath.Join(out, "sheets")
	if err := os.MkdirAll(sheetsDir, 0o755); err != nil {
		log.Fatal(err)
	}

	// seed builds a workbook whose second sheet is named after the given label.
	seed := func(name, value string) []byte {
		wb, err := asposecells.NewWorkbook()
		if err != nil {
			log.Fatal(err)
		}
		empty, err := wb.Save_SaveFormat(asposecells.SaveFormat_Xlsx)
		if err != nil {
			log.Fatal(err)
		}
		bytes, err := editor.EditSpreadsheet(
			datasource.BytesSource(empty),
			editor.InWorksheet(0, editor.SetCellValue(0, 0, name)),
			editor.WithAddWorksheet(name),
			editor.InWorksheet(name, editor.SetCellValue(0, 0, value)),
		)
		if err != nil {
			log.Fatal(err)
		}
		return bytes
	}

	alpha := seed("Alpha", "first")
	beta := seed("Beta", "second")

	// Merge in-memory sources into a single XLSX []byte via a BytesSink.
	var mergedSink datasource.BytesSink
	if err := manipulator.Merge(
		[]datasource.DataSource{datasource.BytesSource(alpha), datasource.BytesSource(beta)},
		ooxml.New(), &mergedSink,
	); err != nil {
		log.Fatalf("merge: %v", err)
	}
	if err := os.WriteFile(examples.OutPath("merge-split", "merged.xlsx"), mergedSink.Bytes(), 0o644); err != nil {
		log.Fatal(err)
	}

	// Merge real sample workbooks from files to a file: FilePathSource inputs,
	// output format from the extension, FilePathSink output.
	mergedFile := examples.OutPath("merge-split", "merged-from-files.xlsx")
	ext := filepath.Ext(mergedFile)[1:]
	opt := formats.Get(ext)
	if opt == nil {
		log.Fatalf("no registered format for %q", ext)
	}
	if err := manipulator.Merge(
		[]datasource.DataSource{
			datasource.FilePathSource(examples.DataPath("CompanySales.xlsx")),
			datasource.FilePathSource(examples.DataPath("EmployeeSalesSummary.xlsx")),
			datasource.FilePathSource(examples.DataPath("BookText.xlsx")),
		},
		opt, datasource.FilePathSink(mergedFile),
	); err != nil {
		log.Fatalf("merge to file: %v", err)
	}

	// Merge straight into a writer (single-file HTML here) via a WriterSink.
	mergedHTML, err := os.Create(examples.OutPath("merge-split", "merged.html"))
	if err != nil {
		log.Fatal(err)
	}
	if err := manipulator.Merge(
		[]datasource.DataSource{datasource.BytesSource(alpha), datasource.BytesSource(beta)},
		html.New(html.WithSaveAsSingleFile(true)), datasource.NewWriterSink(mergedHTML),
	); err != nil {
		log.Fatalf("merge to writer: %v", err)
	}
	mergedHTML.Close()

	// Split the merged workbook back into per-sheet files inside a zip []byte via
	// a ZipSink over an in-memory zip.Writer.
	zipBuf := new(bytes.Buffer)
	zipWriter := zip.NewWriter(zipBuf)
	if err := manipulator.Split(datasource.BytesSource(mergedSink.Bytes()), ooxml.New(), datasource.NewZipSink(zipWriter)); err != nil {
		log.Fatalf("split to zip: %v", err)
	}
	if err := zipWriter.Close(); err != nil {
		log.Fatal(err)
	}
	if err := os.WriteFile(examples.OutPath("merge-split", "split.zip"), zipBuf.Bytes(), 0o644); err != nil {
		log.Fatal(err)
	}

	// Split into a zip.Writer handed in by the caller.
	zipFile, err := os.Create(examples.OutPath("merge-split", "split-writer.zip"))
	if err != nil {
		log.Fatal(err)
	}
	zipWriter2 := zip.NewWriter(zipFile)
	if err := manipulator.Split(datasource.BytesSource(mergedSink.Bytes()), ooxml.New(), datasource.NewZipSink(zipWriter2)); err != nil {
		log.Fatalf("split to zip writer: %v", err)
	}
	if err := zipWriter2.Close(); err != nil {
		log.Fatal(err)
	}
	zipFile.Close()

	// Split the merged file into a folder of standalone files, one per sheet,
	// via a FolderSink. Each file is named <sheet>.<ext>.
	if err := manipulator.Split(datasource.FilePathSource(mergedFile), ooxml.New(), datasource.FolderSink(sheetsDir)); err != nil {
		log.Fatalf("split to folder: %v", err)
	}

	log.Println("merge/split complete; outputs written to examples/merge-split/out")
}

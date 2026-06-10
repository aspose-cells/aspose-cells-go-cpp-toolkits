# manipulator

## Functions

### MergeSpreadsheets

```go
func MergeSpreadsheets 
```

MergeSpreadsheets Merge multi spreadsheets into a file with the specified output format.

Parameters:

  - source: The data sources implementing the datasource.DataSource interface, which provides the input spreadsheets content (e.g., from a file, in-memory buffer, etc.).
  - outSaveOption: Split outfile options that define the output format and behavior, implementing the saveoptions.SaveOption interface (e.g., PDFSaveOption, XLSXSaveOption, CSVSaveOption, etc.).

Returns:

  - \[]byte: The converted file content as a byte slice in the target format.
  - error: An error if the conversion fails due to reasons such as unreadable source, unsupported format, missing license, or failure in the underlying Aspose.Cells engine.

Example:

	save_option = html.New(html.WithExportImagesAsBase64(true), html.WithSaveAsSingleFile(true))
	mergedDataSource := []datasource.DataSource{datasource.FilePathSource("TestData/Source/BookText.xlsx"), datasource.FilePathSource("TestData/Source/EmployeeSalesSummary.xlsx")}
	bytes_data, err = manipulator.MergeSpreadsheets(mergedDataSource, save_option)
	if err != nil {
		println(err)
		return
	}
	os.WriteFile("TestData/Output/mergedOutput2.html", bytes_data, 0644)

### MergeSpreadsheetsToFile

```go
func MergeSpreadsheetsToFile error
```

MergeSpreadsheetsToFile  Merge multi spreadsheets into a file.

Parameters:

  - inputPath: A data source implementing the datasource.DataSource interface, which provides the input spreadsheet content (e.g., from a file, in-memory buffer, HTTP URL, etc.).
  - outputPath: The output file path.

Returns:

  - error: An error if the conversion fails due to reasons such as unreadable source, unsupported format, missing license, or failure in the underlying Aspose.Cells engine.

Example:

manipulator.MergeSpreadsheetsToFile(\[]string{"TestData/Source/CompanySales.xlsx", "TestData/Source/BookText.xlsx", "TestData/Source/EmployeeSalesSummary.xlsx"}, "TestData/Output/MergeBook.xlsx")

### MergeSpreadsheetsToWriter

```go
func MergeSpreadsheetsToWriter error
```

MergeSpreadsheetsToWriter Merge multi spreadsheets into a writer holder with the specified output format.

Parameters:

  - source: The data sources implementing the datasource.DataSource interface, which provides the input spreadsheets content (e.g., from a file, in-memory buffer, etc.).
  - w: An io.Writer (such as a file, bytes.Buffer, or HTTP response writer) where the converted output will be written.
  - outSaveOption: Split outfile options that define the output format and behavior, implementing the saveoptions.SaveOption interface (e.g., PDFSaveOption, XLSXSaveOption, CSVSaveOption, etc.).

Returns:

  - error: An error if the conversion fails due to reasons such as unreadable source, unsupported format, missing license, or failure in the underlying Aspose.Cells engine.

Example:

	save_option = html.New(html.WithExportImagesAsBase64(true), html.WithSaveAsSingleFile(true))
	mergedDataSource := []datasource.DataSource{datasource.FilePathSource("TestData/Source/BookText.xlsx"), datasource.FilePathSource("TestData/Source/EmployeeSalesSummary.xlsx")}
	bytes_data, err = manipulator.MergeSpreadsheets(mergedDataSource, save_option)
	if err != nil {
		println(err)
		return
	}
	os.WriteFile("TestData/Output/mergedOutput2.html", bytes_data, 0644)

### SplitSpreadsheet

```go
func SplitSpreadsheet 
```

### SplitSpreadsheetToFolder

```go
func SplitSpreadsheetToFolder error
```

SplitSpreadsheetToFolder Splits a spreadsheet by worksheet into multiple files in a folder.

Parameters:

  - inputPath: A data source implementing the datasource.DataSource interface, which provides the input spreadsheet content (e.g., from a file, in-memory buffer, HTTP URL, etc.).
  - outputFolder: The output folder which store split result files.

Returns:

  - error: An error if the conversion fails due to reasons such as unreadable source, unsupported format, missing license, or failure in the underlying Aspose.Cells engine.

Example:

manipulator.SplitSpreadsheetToFolder("TestData/Source/BookText.xlsx", "TestData/Output")

### SplitSpreadsheetToZipWriter

```go
func SplitSpreadsheetToZipWriter error
```

SplitSpreadsheetToZipWriter Splits a spreadsheet by worksheet into multiple files in the specified output format. All output files write into a ZIP archive.

Parameters:

  - source: A data source implementing the datasource.DataSource interface, which provides the input spreadsheet content (e.g., from a file, in-memory buffer, HTTP URL, etc.).
  - outSaveOption: Split outfile options that define the output format and behavior, implementing the saveoptions.SaveOption interface (e.g., PDFSaveOption, XLSXSaveOption, CSVSaveOption, etc.).

Returns:

  - error: An error if the conversion fails due to reasons such as unreadable source, unsupported format, missing license, or failure in the underlying Aspose.Cells engine.

Example:

zipWriter := zip.NewWriter(zipFile) save\_option = image.New(image.WithImageType("png")) err = manipulator.SplitSpreadsheetToZipWriter(datasource.FilePathSource("TestData/Source/BookText.xlsx"), zipWriter, save\_option) if err != nil { println(err) return } zipWriter.Flush() zipWriter.Close() zipFile.Close()


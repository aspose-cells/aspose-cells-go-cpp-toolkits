
# Aspose.Cells for Go via C++ Toolkits

## Overview

**Aspose.Cells for Go via C++ Toolkits** is a wrapper toolkit based on [Aspose.Cells for Go via C++](https://products.aspose.com/cells/go-cpp/), designed to provide more Go-idiomatic APIs for Excel document processing. It makes Excel manipulation simpler and more efficient in Go projects.

## Background

While the official Go via C++ version of Aspose.Cells is powerful, its API design style leans towards C++, making it somewhat cumbersome to use in Go projects. This toolkit maintains all the original powerful features while providing:

- **Go-Style APIs**: Function naming and calling conventions that follow Go idioms
- **Optimized Error Handling**: Unified error handling mechanism
- **Common Features Integration**: Encapsulation of frequently used Excel operations

### Supported platforms

- Windows x64
- Linux x64

### Environments and versions

- Go 1.16 or greater
- Aspose.Cells for Go via C++ v26.1.0

## Quick Start

### Create a directory for your project and a main.go file within. Add the following code to your main.go. 

```go
package main

import (
  "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/converter"
  "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/core"
  "os"
)

func main() {
  core.SetLicense(os.Getenv("LicensePath"))
  converter.ConvertFile("TestData/Source/BookText.xlsx", "TestData/Output/output1.pdf")
  converter.ConvertFile("TestData/Source/BookText.xlsx", "TestData/Output/output1.md")
}

```

### Initialize project go.mod

```powershell
go mod init main
```

```
module github.com/aspose-cells/aspose-cells-go-cpp-toolkits

go 1.13

require github.com/aspose-cells/aspose-cells-go-cpp/v26 v26.1.0

```
### Fetch the dependencies for your project.

```powershell
go mod tidy
```
### Set your PATH to point to the shared libraries in Aspose.Cells for Go via C++ in your current command shell.

```powershell
$env:Path = $env:Path+ ";${env:GOPATH}\github.com\aspose-cells\aspose-cells-go-cpp\v26@v26.1.0\lib\win_x86_64\"
```

## Features

### Conversion Spreadsheet

### Support file format

| **Format**                                                        | **Description**                                                                                 | **Load** | **Save** |
| :---------------------------------------------------------------- | :---------------------------------------------------------------------------------------------- | :------- | :------- |
| [XLS](https://docs.fileformat.com/spreadsheet/xls/)               | Excel 95/5.0 - 2003 Workbook.                                                                   | &radic;  | &radic;  |
| [XLSX](https://docs.fileformat.com/spreadsheet/xlsx/)             | Office Open XML SpreadsheetML Workbook or template file, with or without macros.                | &radic;  | &radic;  |
| [XLSB](https://docs.fileformat.com/spreadsheet/xlsb/)             | Excel Binary Workbook.                                                                          | &radic;  | &radic;  |
| [XLSM](https://docs.fileformat.com/spreadsheet/xlsm/)             | Excel Macro-Enabled Workbook.                                                                   | &radic;  | &radic;  |
| [XLT](https://docs.fileformat.com/spreadsheet/xlt/)               | Excel 97 - Excel 2003 Template.                                                                 | &radic;  | &radic;  |
| [XLTX](https://docs.fileformat.com/spreadsheet/xltx/)             | Excel Template.                                                                                 | &radic;  | &radic;  |
| [XLTM](https://docs.fileformat.com/spreadsheet/xltm/)             | Excel Macro-Enabled Template.                                                                   | &radic;  | &radic;  |
| [XLAM](https://docs.fileformat.com/spreadsheet/xlam/)             | An Excel Macro-Enabled Add-In file that's used to add new functions to Excel.                   |          | &radic;  |
| [CSV](https://docs.fileformat.com/spreadsheet/csv/)               | CSV (Comma Separated Value) file.                                                               | &radic;  | &radic;  |
| [TSV](https://docs.fileformat.com/spreadsheet/tsv/)               | TSV (Tab-separated values) file.                                                                | &radic;  | &radic;  |
| [TXT](https://docs.fileformat.com/word-processing/txt/)           | Delimited plain text file.                                                                      | &radic;  | &radic;  |
| [HTML](https://docs.fileformat.com/web/html/)                     | HTML format.                                                                                    | &radic;  | &radic;  |
| [MHTML](https://docs.fileformat.com/web/mhtml/)                   | MHTML file.                                                                                     | &radic;  | &radic;  |
| [ODS](https://docs.fileformat.com/spreadsheet/ods/)               | ODS (OpenDocument Spreadsheet).                                                                 | &radic;  | &radic;  |
| [JSON](https://docs.fileformat.com/web/json/)                     | JavaScript Object Notation                                                                      | &radic;  | &radic;  |
| [DIF](https://docs.fileformat.com/spreadsheet/dif/)               | Data Interchange Format.                                                                        |          | &radic;  |
| [PDF](https://docs.fileformat.com/pdf/)                           | Adobe Portable Document Format.                                                                 |          | &radic;  |
| [XPS](https://docs.fileformat.com/page-description-language/xps/) | XML Paper Specification Format.                                                                 |          | &radic;  |
| [SVG](https://docs.fileformat.com/page-description-language/svg/) | Scalable Vector Graphics Format.                                                                |          | &radic;  |
| [TIFF](https://docs.fileformat.com/image/tiff/)                   | Tagged Image File Format                                                                        |          | &radic;  |
| [PNG](https://docs.fileformat.com/image/png/)                     | Portable Network Graphics Format                                                                |          | &radic;  |
| [BMP](https://docs.fileformat.com/image/bmp/)                     | Bitmap Image Format                                                                             |          | &radic;  |
| [EMF](https://docs.fileformat.com/image/emf/)                     | Enhanced metafile Format                                                                        |          | &radic;  |
| [JPEG](https://docs.fileformat.com/image/jpeg/)                   | JPEG is a type of image format that is saved using the method of lossy compression.             |          | &radic;  |
| [GIF](https://docs.fileformat.com/image/gif/)                     | Graphical Interchange Format                                                                    |          | &radic;  |
| [MARKDOWN](https://docs.fileformat.com/word-processing/md/)       | Represents a markdown document.                                                                 |          | &radic;  |
| [SXC](https://docs.fileformat.com/spreadsheet/sxc/)               | An XML based format used by OpenOffice and StarOffice                                           | &radic;  | &radic;  |
| [FODS](https://docs.fileformat.com/spreadsheet/fods/)             | This is an Open Document format stored as flat XML.                                             | &radic;  | &radic;  |
| [DOCX](https://docs.fileformat.com/word-processing/docx/)         | A well-known format for Microsoft Word documents that is a combination of XML and binary files. |          | &radic;  |
| [PPTX](https://docs.fileformat.com/presentation/pptx/)            | The PPTX format is based on the Microsoft PowerPoint open XML presentation file format.         |          | &radic;  |

## Evaluate Aspose.Cells for Go via C++ Toolkits

You can use Aspose.Cells for Go via C++ Toolkits free of cost for evaluation.The evaluation version provides almost all functionality of the product with certain limitations. The same evaluation version becomes licensed when you purchase a license and add a couple of lines of code to apply the license.
If you want to test Aspose.Cells for Go via C++ without evaluation version limitations, you can also try a 30-Day Temporary License. Please refer to <a href="https://purchase.aspose.com/temporary-license/"> How to get a Temporary License</a>?

## Limitations of Evaluation version

The evaluation version of Aspose. Cells for Go Toolset provides complete product functionality. An evaluation watermark will be inserted when saving the file. And the evaluation version can open up to 100 files.

## Run Aspose.Cells for Go via C++ Toolkits in production

A commercial license key is required in a production environment. Please contact us to <a href="https://purchase.aspose.com/buy">purchase a commercial license</a> if you want to publish application to the product server.


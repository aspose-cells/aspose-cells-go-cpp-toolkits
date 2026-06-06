package editor

import asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"

type WorkbookAction func(workbook *asposecells.Workbook) error

type WorksheetAction func(worksheet *asposecells.Worksheet) error

//type ChartAction func(chart *asposecells.Chart) error

package cells

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"io"
)

func GetWorkbookWithDataSource(source datasource.DataSource) (*asposecells.Workbook, error) {
	reader, errOpen := source.Open()
	if errOpen != nil {
		return nil, errOpen
	}
	data, errRead := io.ReadAll(reader)
	if errRead != nil {
		return nil, errRead
	}
	reader.Close()
	return asposecells.NewWorkbook_Stream(data)
}

func WorkbookToByteData(workbook *asposecells.Workbook) ([]byte, error) {
	fileFormat, err := workbook.GetFileFormat()
	if err != nil {
		return nil, err
	}
	saveFormat := formats.FileFormatToSaveFormat(fileFormat)
	return workbook.Save_SaveFormat(saveFormat)
}

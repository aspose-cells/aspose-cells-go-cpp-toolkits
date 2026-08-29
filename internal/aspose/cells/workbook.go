package cells

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"io"
)

// ReadSource reads all bytes from a datasource.DataSource. It opens the source,
// drains it fully into memory, and closes it. This is the single place where a
// DataSource is turned into raw bytes, so every package reads sources the same way.
func ReadSource(source datasource.DataSource) ([]byte, error) {
	reader, errOpen := source.Open()
	if errOpen != nil {
		return nil, errOpen
	}
	defer reader.Close()
	data, errRead := io.ReadAll(reader)
	if errRead != nil {
		return nil, errRead
	}
	return data, nil
}

func GetWorkbookWithDataSource(source datasource.DataSource) (*asposecells.Workbook, error) {
	data, err := ReadSource(source)
	if err != nil {
		return nil, err
	}
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

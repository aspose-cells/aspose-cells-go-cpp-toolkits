package main

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/converter"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/core"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/markdown"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/pdf"
	"os"
)

func main() {
	core.SetLicense(os.Getenv("LicensePath"))
	converter.ConvertFile("TestData/Source/BookText.xlsx", "TestData/Output/output1.pdf")
	converter.ConvertFile("TestData/Source/BookText.xlsx", "TestData/Output/output1.md")
	save_option := pdf.New(pdf.WithOnePagePerSheet(true))
	bytes_data, err := converter.ConvertSpreadsheet(datasource.FilePathSource("TestData/Source/BookText.xlsx"), save_option)
	if err != nil {
		println(err)
		return
	}
	os.WriteFile("TestData/Output/output2.pdf", bytes_data, 0644)

	file, err := os.Create("TestData/Output/output2.md")
	if err != nil {
		panic(err)
	}
	err = converter.ConvertToWriter(datasource.FilePathSource("TestData/Source/BookText.xlsx"), markdown.New(markdown.WithClearData(true)), file)
	if err != nil {
		return
	}
	defer file.Close()

	converter.ConvertFile("TestData/Source/BookText.xlsx", "TestData/Output/output2.xps")
	converter.ConvertFile("TestData/Source/BookText.xlsx", "TestData/Output/output2.tif")
	converter.ConvertFile("TestData/Source/BookText.xlsx", "TestData/Output/output2.svg")
	converter.ConvertFile("TestData/Source/BookText.xlsx", "TestData/Output/output2.json")
}

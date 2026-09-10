package tests

import (
	"errors"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/transfer"
)

// TestTransferImportJsonData verifies the ImportJsonData function works correctly.
func TestTransferImportJsonData(t *testing.T) {
	// Create a seed workbook
	seedBytes := newTestWorkbookBytes(t)

	// JSON data to import
	jsonData := []byte(`{
		"values": [
			{"name": "Alice", "age": 30},
			{"name": "Bob", "age": 25}
		]
	}`)

	// Import JSON data
	var out datasource.BytesSink
	err := transfer.ImportJsonData(
		datasource.BytesSource(seedBytes),
		datasource.BytesSource(jsonData),
		&out,
		transfer.WithSheet("Sheet1"),
		transfer.WithBeginCell(0, 0),
	)
	if err != nil {
		t.Fatalf("ImportJsonData error: %v", err)
	}

	// Verify the workbook was modified
	result := out.Bytes()
	if len(result) == 0 {
		t.Fatal("ImportJsonData returned empty result")
	}
}

// TestTransferImportXmlData verifies the ImportXmlData function works correctly.
func TestTransferImportXmlData(t *testing.T) {
	// Create a seed workbook
	seedBytes := newTestWorkbookBytes(t)

	// XML data to import
	xmlData := []byte(`<?xml version="1.0" encoding="utf-8"?>
<root>
	<row>
		<name>Alice</name>
		<age>30</age>
	</row>
	<row>
		<name>Bob</name>
		<age>25</age>
	</row>
</root>`)

	// Import XML data
	var out datasource.BytesSink
	err := transfer.ImportXMLData(
		datasource.BytesSource(seedBytes),
		datasource.BytesSource(xmlData),
		&out,
		transfer.WithSheet("Sheet1"),
		transfer.WithBeginCell(0, 0),
	)
	if err != nil {
		t.Fatalf("ImportXMLData error: %v", err)
	}

	// Verify the workbook was modified
	result := out.Bytes()
	if len(result) == 0 {
		t.Fatal("ImportXMLData returned empty result")
	}
}

// TestTransferImportJsonDataNilSource verifies ErrDataSourceNil is returned when source is nil.
func TestTransferImportJsonDataNilSource(t *testing.T) {
	var out datasource.BytesSink
	err := transfer.ImportJsonData(nil, datasource.BytesSource([]byte{}), &out)
	if !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Fatalf("ImportJsonData(nil, ...) error = %v, want ErrDataSourceNil", err)
	}
}

// TestTransferImportJsonDataNilJsonData verifies ErrDataSourceNil is returned when jsonData is nil.
func TestTransferImportJsonDataNilJsonData(t *testing.T) {
	seedBytes := newTestWorkbookBytes(t)
	var out datasource.BytesSink
	err := transfer.ImportJsonData(datasource.BytesSource(seedBytes), nil, &out)
	if !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Fatalf("ImportJsonData(..., nil) error = %v, want ErrDataSourceNil", err)
	}
}

// TestTransferImportJsonDataNilSink verifies ErrDataSinkNil is returned when sink is nil.
func TestTransferImportJsonDataNilSink(t *testing.T) {
	seedBytes := newTestWorkbookBytes(t)
	err := transfer.ImportJsonData(datasource.BytesSource(seedBytes), datasource.BytesSource([]byte{}), nil)
	if !errors.Is(err, toolkiterrors.ErrDataSinkNil) {
		t.Fatalf("ImportJsonData(..., nil) error = %v, want ErrDataSinkNil", err)
	}
}

// TestTransferImportXmlDataNilSource verifies ErrDataSourceNil is returned when source is nil.
func TestTransferImportXmlDataNilSource(t *testing.T) {
	var out datasource.BytesSink
	err := transfer.ImportXMLData(nil, datasource.BytesSource([]byte{}), &out)
	if !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Fatalf("ImportXMLData(nil, ...) error = %v, want ErrDataSourceNil", err)
	}
}

// TestTransferImportXmlDataNilXmlData verifies ErrDataSourceNil is returned when xmlData is nil.
func TestTransferImportXmlDataNilXmlData(t *testing.T) {
	seedBytes := newTestWorkbookBytes(t)
	var out datasource.BytesSink
	err := transfer.ImportXMLData(datasource.BytesSource(seedBytes), nil, &out)
	if !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
		t.Fatalf("ImportXMLData(..., nil) error = %v, want ErrDataSourceNil", err)
	}
}

// TestTransferImportXmlDataNilSink verifies ErrDataSinkNil is returned when sink is nil.
func TestTransferImportXmlDataNilSink(t *testing.T) {
	seedBytes := newTestWorkbookBytes(t)
	err := transfer.ImportXMLData(datasource.BytesSource(seedBytes), datasource.BytesSource([]byte{}), nil)
	if !errors.Is(err, toolkiterrors.ErrDataSinkNil) {
		t.Fatalf("ImportXMLData(..., nil) error = %v, want ErrDataSinkNil", err)
	}
}

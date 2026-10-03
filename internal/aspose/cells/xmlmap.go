package cells

import (
	"bytes"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"

	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// LooksLikeXML reports whether data is an XML document rather than one of the
// engine's own binary formats.
//
// The engine loads an XML document in two very different ways: as a plain
// workbook — which discards the schema and leaves the workbook with no XML map
// at all — or with XmlLoadOptions.IsXmlMap, which keeps the schema as a
// workbook XML map and the element data as worksheet cells. Only the second
// makes the workbook re-exportable as XML, so the loader has to tell the two
// apart before it decides which call to make.
//
// Detection is a sniff of the leading bytes, not a parse: skip a UTF-8 BOM and
// leading whitespace and require '<'. Every format the engine writes itself
// (xlsx, xlsb, ods) is a zip and so begins "PK", and the legacy binary formats
// begin with their own magic numbers, so none of them can start with '<'.
func LooksLikeXML(data []byte) bool {
	data = bytes.TrimPrefix(data, []byte{0xEF, 0xBB, 0xBF})
	data = bytes.TrimLeft(data, " \t\r\n")
	return len(data) > 0 && data[0] == '<'
}

// XMLMapNames returns the names of the workbook's XML maps, in collection
// order. A workbook with no XML maps yields an empty slice.
func XMLMapNames(workbook *asposecells.Workbook) ([]string, error) {
	worksheets, err := engine.Derive(workbook.GetWorksheets())
	if err != nil {
		return nil, err
	}
	maps, err := worksheets.GetXmlMaps()
	if err != nil {
		return nil, err
	}
	count, err := maps.GetCount()
	if err != nil {
		return nil, err
	}
	names := make([]string, 0, count)
	for i := int32(0); i < count; i++ {
		xmlMap, err := maps.Get(i)
		if err != nil {
			return nil, err
		}
		name, err := xmlMap.GetName()
		if err != nil {
			return nil, err
		}
		names = append(names, name)
	}
	return names, nil
}

// GetWorkbookFromBytes loads an in-memory workbook from bytes.
func GetWorkbookFromBytes(data []byte) (*asposecells.Workbook, error) {
	return engine.OpenWorkbook(data)
}

// GetWorkbookFromBytesWithXMLLoad loads an in-memory workbook from XML bytes,
// keeping the document's schema as a workbook XML map (see LooksLikeXML). The
// engine names the map after the document's root element — "Rows_Map" for a
// root element of "Rows" — so callers read the name back with XMLMapNames
// rather than predicting it.
func GetWorkbookFromBytesWithXMLLoad(data []byte) (*asposecells.Workbook, error) {
	xmlOptions, err := asposecells.NewXmlLoadOptions()
	if err != nil {
		return nil, err
	}
	engine.Disarm(xmlOptions)
	if err := xmlOptions.SetIsXmlMap(true); err != nil {
		return nil, err
	}
	return engine.OpenWorkbookWithOptions(data, xmlOptions.ToLoadOptions())
}

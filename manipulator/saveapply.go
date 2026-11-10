package manipulator

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// saveWorkbookBytes writes workbook through opt.
//
// When opt implements saveoptions.WorkbookApplier the workbook is saved
// directly, which avoids serializing it and reloading a second workbook just
// to save it. Third-party SaveOption implementations that only provide Apply
// keep working through the serialize/reload fallback.
func saveWorkbookBytes(opt saveoptions.SaveOption, workbook *asposecells.Workbook) ([]byte, error) {
	if applier, ok := opt.(saveoptions.WorkbookApplier); ok {
		return applier.SaveWorkbook(workbook)
	}
	data, err := workbook.SaveToStream()
	if err != nil {
		return nil, err
	}
	return opt.Apply(data)
}

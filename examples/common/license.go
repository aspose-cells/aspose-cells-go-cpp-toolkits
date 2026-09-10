package examples

import (
	"os"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/core"
)

// SetLicense applies the Aspose.Cells license from the LicenseFilePath
// environment variable. When the variable is empty it falls back to the legacy
// LicensePath variable, and when that is empty too it does nothing and leaves
// the engine in evaluation mode. The returned error (wrapping
// errors.ErrLicenseInvalid) is non-nil only when a path was configured but the
// license could not be applied, so callers can log and continue.
func SetLicense() error {
	path := os.Getenv("LicenseFilePath")
	if path == "" {
		path = os.Getenv("LicensePath")
	}
	if path == "" {
		return nil
	}
	return core.SetLicense(path)
}

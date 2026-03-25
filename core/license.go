package core

import (
	"os"

	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// SetLicense loads and applies a license for the Aspose.Cells engine from the specified file path.
//
// Parameters:
//   - licensePath: The absolute or relative path to a valid Aspose.Cells license file (typically with a .lic extension).
//     If the file is not found, unreadable, or contains an invalid license, subsequent operations may run in
//     evaluation mode (e.g., with watermarks or feature limitations).
//
// Notes:
//   - This function configures the global license state for the underlying Aspose.Cells for C++ library.
//     It should be called once at application startup before performing any conversion or manipulation operations.
//   - Calling SetLicense multiple times has no effect after a valid license is successfully loaded.
//   - If no license is set, the library operates in trial mode, which may impose restrictions such as:
//   - Watermarks on output documents,
//   - Limited worksheet size,
//   - The license file must be obtained from Aspose and is bound to your subscription or purchase.
//
// Example:
//
//	core.SetLicense(os.Getenv("LicensePath"))
func SetLicense(licensePath string) {
	if licensePath == "" {
		licensePath = os.Getenv("LicensePath")
	}
	lic, _ := asposecells.NewLicense()
	err := lic.SetLicense_String(licensePath)
	if err != nil {
		println(err)
	}
}

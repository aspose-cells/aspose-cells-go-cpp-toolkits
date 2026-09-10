// Package core holds engine-level configuration.
//
// Currently it exposes SetLicense for licensing the underlying Aspose.Cells
// engine before any spreadsheet is processed.
package core

import (
	"fmt"
	"os"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// SetLicense loads and applies a license for the Aspose.Cells engine from the
// specified file path.
//
// Parameters:
//   - licensePath: The absolute or relative path to a valid Aspose.Cells license
//     file (typically with a .lic extension). If empty, the LicenseFilePath
//     environment variable is used (falling back to the legacy LicensePath
//     variable). If the file is not found or unreadable, SetLicense reports it
//     directly; if it is readable but contains an invalid license, subsequent
//     operations may run in evaluation mode (e.g., with watermarks or feature
//     limitations).
//
// Returns:
//   - error: nil when the license was applied; otherwise an error wrapping
//     ErrLicenseInvalid (plus the underlying engine error). Callers that need
//     to detect license failure should check the return value; a statement
//     call without one keeps compiling and simply ignores the failure.
//
// Notes:
//   - This function configures the global license state for the underlying
//     Aspose.Cells for C++ library. It should be called once at application
//     startup before performing any conversion or manipulation operations.
//   - Calling SetLicense multiple times has no effect after a valid license is
//     successfully loaded.
//   - If no license is set, the library operates in trial mode, which may
//     impose restrictions such as watermarks on output documents or a limited
//     worksheet size.
//
// Example:
//
//	if err := core.SetLicense(os.Getenv("LicenseFilePath")); err != nil {
//		log.Fatal(err)
//	}
func SetLicense(licensePath string) error {
	if licensePath == "" {
		licensePath = os.Getenv("LicenseFilePath")
	}
	if licensePath == "" {
		// Legacy name, kept for callers that set it instead.
		licensePath = os.Getenv("LicensePath")
	}
	if licensePath != "" {
		// The engine reports a missing/unreadable license file silently and
		// just keeps running in evaluation mode, so check readability here to
		// honour the documented ErrLicenseInvalid contract.
		if _, err := os.Stat(licensePath); err != nil {
			return fmt.Errorf("license file %q: %w: %v", licensePath, toolkiterrors.ErrLicenseInvalid, err)
		}
	}
	lic, err := asposecells.NewLicense()
	if err != nil {
		return fmt.Errorf("create license: %w: %v", toolkiterrors.ErrLicenseInvalid, err)
	}
	if err := lic.SetLicense_String(licensePath); err != nil {
		return fmt.Errorf("apply license %q: %w: %v", licensePath, toolkiterrors.ErrLicenseInvalid, err)
	}
	return nil
}

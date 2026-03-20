package core

import (
	"os"

	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

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

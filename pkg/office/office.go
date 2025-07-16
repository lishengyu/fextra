package office

import (
	"fextra/internal"
	"fextra/pkg/office/doc"
)

func init() {
	// doc(7)
	internal.RegisterParser(internal.FileTypeDOC, &doc.OfficeDocParser{})
}

package office

import (
	"fextra/internal"
	"fextra/pkg/office/doc"
	"fextra/pkg/office/docx"
	"fextra/pkg/office/ppt"
	"fextra/pkg/office/pptx"
	"fextra/pkg/office/xlsx"
)

func init() {
	// doc(7)
	internal.RegisterParser(internal.FileTypeDOC, &doc.OfficeDocParser{})
	internal.RegisterParser(internal.FileTypePPT, &ppt.OfficePptParser{})
	internal.RegisterParser(internal.FileTypeDOCX, &docx.OfficeDocxParser{})
	internal.RegisterParser(internal.FileTypePPTX, &pptx.OfficePptxParser{})
	internal.RegisterParser(internal.FileTypeXLSX, &xlsx.OfficeXlsxParser{})
}

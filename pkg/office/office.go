package office

import (
	"fextra/internal"
	"fextra/pkg/office/doc"
	"fextra/pkg/office/docx"
	"fextra/pkg/office/odt"
	"fextra/pkg/office/pdf"
	"fextra/pkg/office/ppt"
	"fextra/pkg/office/pptx"
	"fextra/pkg/office/rtf"
	"fextra/pkg/office/vsd"
	"fextra/pkg/office/vsdx"
	"fextra/pkg/office/xlsb"
	"fextra/pkg/office/xlsx"
)

func init() {
	// doc(7)
	internal.RegisterParser(internal.FileTypeDOC, &doc.OfficeDocParser{})
	internal.RegisterParser(internal.FileTypePPT, &ppt.OfficePptParser{})
	internal.RegisterParser(internal.FileTypeDOCX, &docx.OfficeDocxParser{})
	internal.RegisterParser(internal.FileTypePPTX, &pptx.OfficePptxParser{})
	internal.RegisterParser(internal.FileTypeXLSX, &xlsx.OfficeXlsxParser{})
	internal.RegisterParser(internal.FileTypeRTF, &rtf.OfficeRtfParser{})
	internal.RegisterParser(internal.FileTypeODT, &odt.OfficeOdtParser{})
	internal.RegisterParser(internal.FileTypePDF, &pdf.OfficePdfParser{})
	internal.RegisterParser(internal.FileTypeVSDX, &vsdx.OfficeVsdxParser{})
	internal.RegisterParser(internal.FileTypeXLSB, &xlsb.OfficeXlsbParser{})
	internal.RegisterParser(internal.FileTypeVSD, &vsd.OfficeVsdParser{})
}

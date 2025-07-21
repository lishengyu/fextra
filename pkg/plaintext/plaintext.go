package plaintext

import (
	"fextra/internal"
	"fextra/pkg/plaintext/plainhtml"
	"fextra/pkg/plaintext/plainmd"
	"fextra/pkg/plaintext/plaintxt"
	"fextra/pkg/plaintext/plainxml"
)

func init() {
	// html:1 txt:2  xml:3  json:4   csv:5
	internal.RegisterParser(internal.FileTypeTXT, &plaintxt.TextPlainParser{})
	internal.RegisterParser(internal.FileTypeCSV, &plaintxt.TextPlainParser{})
	internal.RegisterParser(internal.FileTypeXML, &plainxml.TextXMLParser{})
	internal.RegisterParser(internal.FileTypeJSON, &plaintxt.TextPlainParser{})
	internal.RegisterParser(internal.FileTypeHTML, &plainhtml.TextHTMLParser{})
	internal.RegisterParser(internal.FileTypeMD, &plainmd.TextMarkdownParser{})
}

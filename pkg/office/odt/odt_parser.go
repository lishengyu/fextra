package odt

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fextra/pkg/logger"
	"fmt"
	"io"
)

// OfficeOdtParser ODT文档解析器
type OfficeOdtParser struct{}

// Parse 解析ODT文件并提取文本内容
func (p *OfficeOdtParser) Parse(filePath string) ([]byte, error) {
	// 打开ODT文件（ZIP格式）
	zipReader, err := zip.OpenReader(filePath)
	if err != nil {
		return []byte{}, fmt.Errorf("无法打开ODT文件: %v", err)
	}
	defer zipReader.Close()

	// 查找content.xml文件
	var contentFile *zip.File
	for _, file := range zipReader.File {
		if file.Name == "content.xml" {
			contentFile = file
			break
		}
	}

	if contentFile == nil {
		return []byte{}, fmt.Errorf("content.xml不存在于ODT文件中")
	}

	// 读取content.xml内容
	xmlFile, err := contentFile.Open()
	if err != nil {
		return []byte{}, err
	}
	defer xmlFile.Close()

	// 解析XML并提取文本内容
	var textBuilder bytes.Buffer
	var inTextElement bool
	odtTextNS := "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
	d := xml.NewDecoder(xmlFile)

	for {
		token, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			logger.Logger.Printf("XML解析错误: %v", err)
			continue
		}

		switch t := token.(type) {
		case xml.StartElement:
			// 检测文本段落元素
			if t.Name.Space == odtTextNS && (t.Name.Local == "p" || t.Name.Local == "h" || t.Name.Local == "span") {
				inTextElement = true
			}
		case xml.EndElement:
			// 结束文本段落元素
			if t.Name.Space == odtTextNS && (t.Name.Local == "p" || t.Name.Local == "h") {
				inTextElement = false
				textBuilder.WriteString("\n") // 段落结束添加换行
			} else if t.Name.Space == odtTextNS && t.Name.Local == "span" {
				inTextElement = false
			}
		case xml.CharData:
			// 仅收集文本元素内的内容
			if inTextElement {
				textBuilder.WriteString(string(t))
			}
		}
	}

	return textBuilder.Bytes(), nil
}

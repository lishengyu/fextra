package docx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"strings"

	"fextra/pkg/logger"
)

type OfficeDocxParser struct{}

// Parse 提取DOCX文件中的文本内容
func (p *OfficeDocxParser) Parse(filename string) ([]byte, error) {
	// 打开DOCX文件（ZIP格式）
	zipReader, err := zip.OpenReader(filename)
	if err != nil {
		return nil, fmt.Errorf("无法打开DOCX文件: %w", err)
	}
	defer zipReader.Close()

	// 查找word/document.xml文件
	docFile, err := findDocumentXml(zipReader.File)
	if err != nil {
		return nil, fmt.Errorf("找不到document.xml: %w", err)
	}

	// 读取XML内容
	xmlContent, err := readZipFile(docFile)
	if err != nil {
		return nil, fmt.Errorf("无法读取XML内容: %w", err)
	}

	// 解析XML提取文本
	extractedText, err := parseDocumentXml(xmlContent)
	if err != nil {
		return nil, fmt.Errorf("解析XML失败: %w", err)
	}

	return extractedText, nil
}

// findDocumentXml 在ZIP文件中查找word/document.xml
func findDocumentXml(files []*zip.File) (*zip.File, error) {
	for _, file := range files {
		logger.Logger.Printf("docx 文件: %s", file.Name)
		if file.Name == "word/document.xml" {
			return file, nil
		}
	}
	return nil, errors.New("word/document.xml not found in DOCX file")
}

// readZipFile 读取ZIP文件内容
func readZipFile(zf *zip.File) ([]byte, error) {
	rc, err := zf.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	return io.ReadAll(rc)
}

// documentXml 用于解析document.xml的结构
type documentXml struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main document"`
	Body    body     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main body"`
}

type body struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main body"`
	Paras   []para   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main p"` // 段落
}

// 定义WML命名空间常量
const wNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

type para struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main p"`
	PStyle  pStyle   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main pPr>http://schemas.openxmlformats.org/wordprocessingml/2006/main pStyle"` // 段落样式
	Runs    []run    `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main r"`                                                                       // 文本 run
}

type pStyle struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"` // 样式值，如 Heading1, Heading2
}

type run struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main r"`
	Texts   []text   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main t"` // 文本内容
}

type text struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main t"`
	Value   string   `xml:",chardata"`
}

// parseDocumentXml 解析XML内容并提取文本
func parseDocumentXml(xmlContent []byte) ([]byte, error) {
	var doc documentXml
	if err := xml.Unmarshal(xmlContent, &doc); err != nil {
		return []byte{}, err
	}

	var textBuffer bytes.Buffer
	for _, para := range doc.Body.Paras {
		style := para.PStyle.Val
		var paraText bytes.Buffer
		// 提取段落文本内容
		for _, run := range para.Runs {
			for _, t := range run.Texts {
				paraText.WriteString(t.Value)
			}
		}
		// 根据样式添加标识
		if strings.HasPrefix(style, "Heading") {
			textBuffer.WriteString(fmt.Sprintf("【标题%s】 ", style[7:]))
		}
		textBuffer.WriteString(paraText.String())
		textBuffer.WriteString("\n") // 段落间添加换行
	}

	return textBuffer.Bytes(), nil
}

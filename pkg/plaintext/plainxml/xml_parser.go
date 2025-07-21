package plainxml

import (
	"bytes"
	"encoding/xml"
	"fextra/pkg/logger"
	"fmt"
	"html"
	"io"
	"os"
	"regexp"
	"strings"
)

type TextXMLParser struct{}

// TextXMLParser 用于解析XML并提取纯文本内容
var (
	invisibleCharsRegex *regexp.Regexp
	whitespaceRegex     *regexp.Regexp
)

// NewXMLParser 创建XMLParser实例并预编译正则表达式
func init() {
	invisibleCharsRegex = regexp.MustCompile(`[\x00-\x08\x0B\x0C\x0E-\x1F\x7F\x{200B}\x{200C}\x{200D}\x{200E}\x{200F}\x{2028}\x{2029}\x{FEFF}]`)
	whitespaceRegex = regexp.MustCompile(`[\s\x{A0}\x{2000}-\x{200A}\x{2028}\x{2029}\x{202F}\x{205F}\x{3000}]+`)
}

// Parse 从XML内容中提取纯文本
func (p *TextXMLParser) ParseXml(xmlContent []byte) ([]byte, error) {
	decoder := xml.NewDecoder(bytes.NewReader(xmlContent))
	decoder.Strict = false                // 容忍格式不严格的XML
	decoder.AutoClose = xml.HTMLAutoClose // 自动关闭常见标签

	var textSegments []string
	depth := 0

	for {
		token, err := decoder.Token()
		if err != nil {
			if err.Error() == io.EOF.Error() {
				break
			}
			return nil, fmt.Errorf("xml decode error: %w", err)
		}

		switch t := token.(type) {
		case xml.CharData:
			// 处理文本节点
			text := strings.TrimSpace(string(t))
			if text != "" {
				textSegments = append(textSegments, text)
			}
			logger.Logger.Printf("text: %s", text)
		case xml.StartElement:
			depth++
			logger.Logger.Printf("depth: %d, start element: %v", depth, t)
		case xml.EndElement:
			if depth > 0 {
				depth--
			}

		// 忽略注释、指令、处理指令和CDATA节点
		case xml.Comment, xml.Directive, xml.ProcInst:
			continue
		}
	}

	// 处理提取到的文本
	text := strings.Join(textSegments, " ")
	text = html.UnescapeString(text)
	text = invisibleCharsRegex.ReplaceAllString(text, "")
	text = whitespaceRegex.ReplaceAllString(text, " ")

	return []byte(strings.TrimSpace(text)), nil
}

// ParseFile 从XML文件中提取纯文本
func (p *TextXMLParser) Parse(filePath string) ([]byte, error) {
	content, err := os.ReadFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("read xml file error: %w", err)
	}

	return p.ParseXml(content)
}

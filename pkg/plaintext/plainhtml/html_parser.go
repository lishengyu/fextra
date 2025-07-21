package plainhtml

import (
	"bytes"
	"fmt"
	"regexp"
	"strings"

	"os"

	"golang.org/x/net/html"
)

type TextHTMLParser struct{}

// TextHTMLParser 用于解析HTML并提取可视化文本内容
var (
	invisibleCharsRegex *regexp.Regexp
	newlineRegex        *regexp.Regexp
	whitespaceRegex     *regexp.Regexp
)

// NewTextHTMLParser 创建TextHTMLParser实例并预编译正则表达式
func init() {
	invisibleCharsRegex = regexp.MustCompile(`[\x00-\x08\x0B\x0C\x0E-\x1F\x7F\x{200B}\x{200C}\x{200D}\x{200E}\x{200F}\x{2028}\x{2029}\x{FEFF}]`)
	newlineRegex = regexp.MustCompile(`\n+`)
	whitespaceRegex = regexp.MustCompile(`[\s\x{A0}\x{2000}-\x{200A}\x{2028}\x{2029}\x{202F}\x{205F}\x{3000}]+`)
}

// Parse 从HTML内容中提取可视化文本，剥离标签和不可见字符
func (p *TextHTMLParser) ParseHtml(htmlContent []byte) ([]byte, error) {
	// 解析HTML
	doc, err := html.Parse(bytes.NewReader(htmlContent))
	if err != nil {
		return []byte{}, fmt.Errorf("html parse error: %w", err)
	}

	// 提取文本内容
	var textSegments []string
	var extractText func(*html.Node)

	extractText = func(n *html.Node) {
		// 仅处理文本节点
		if n.Type == html.TextNode {
			// 只添加非空白文本
			trimmedText := strings.TrimSpace(n.Data)
			if trimmedText != "" {
				textSegments = append(textSegments, trimmedText)
			}
			return
		}

		// 忽略脚本、样式、头部和元数据标签内容
		if n.Type == html.ElementNode {
			if n.Data == "script" || n.Data == "style" || n.Data == "head" || n.Data == "meta" || n.Data == "link" {
				return
			}
			// 特别处理br标签为空格
			if n.Data == "br" {
				textSegments = append(textSegments, " ")
			}
		}

		// 递归处理子节点
		for c := n.FirstChild; c != nil; c = c.NextSibling {
			extractText(c)
		}
	}

	extractText(doc)

	extractedText := p.processExtractedText(strings.Join(textSegments, " "))
	return []byte(extractedText), nil
}

// ParseFile 从HTML文件中提取可视化文本
// processExtractedText 处理提取到的文本：去除HTML实体、过滤不可见字符、规范化空白
func (p *TextHTMLParser) processExtractedText(rawText string) string {
	extractedText := html.UnescapeString(rawText)
	extractedText = invisibleCharsRegex.ReplaceAllString(extractedText, "")
	extractedText = newlineRegex.ReplaceAllString(extractedText, " ")
	extractedText = whitespaceRegex.ReplaceAllString(extractedText, " ")
	return strings.TrimSpace(extractedText)
}

func (p *TextHTMLParser) Parse(filePath string) ([]byte, error) {
	// 读取文件内容
	fileContent, err := os.ReadFile(filePath)
	if err != nil {
		return []byte{}, fmt.Errorf("读取HTML文件 '%s' 失败: %w", filePath, err)
	}

	return p.ParseHtml(fileContent)
}

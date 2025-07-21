package plainmd

import (
	"fextra/pkg/logger"
	"os"
	"regexp"
	"strings"

	"fmt"

	"github.com/yuin/goldmark"
	"github.com/yuin/goldmark/ast"
	"github.com/yuin/goldmark/text"
)

type TextMarkdownParser struct{}

// MarkdownParser 用于提取Markdown文本内容的解析器
var (
	invisibleCharsRegex *regexp.Regexp
	whitespaceRegex     *regexp.Regexp
	newlineRegex        *regexp.Regexp
)

func init() {
	invisibleCharsRegex = regexp.MustCompile(`[\x00-\x08\x0B\x0C\x0E-\x1F\x7F\x{200B}-\x{200F}\x{FEFF}]`)

	// 是否保留换行符，通过调整正则表达式来实现
	// whitespaceRegex = regexp.MustCompile(`[\s\x{A0}\x{2000}-\x{200A}\x{2028}\x{2029}\x{3000}]+`)
	whitespaceRegex = regexp.MustCompile(`[\t\f\v\x{A0}\x{2000}-\x{200A}\x{2028}\x{2029}\x{3000}]+`)
	newlineRegex = regexp.MustCompile(`\n+`)
}

// Parse 从Markdown字节内容中提取纯文本
func (p *TextMarkdownParser) ParseMd(content []byte) (string, error) {
	md := goldmark.New()
	reader := text.NewReader(content)
	rootNode := md.Parser().Parse(reader) // 生成 AST 根节点

	var textSegments []string
	ast.Walk(rootNode, func(node ast.Node, entering bool) (ast.WalkStatus, error) {
		if entering {
			logger.DebugLogger.Printf("Node Kind: %s", node.Kind())
			switch n := node.(type) {
			case *ast.Text:
				logger.DebugLogger.Printf("Text: %s", string(n.Value(content)))
				textSegments = append(textSegments, string(n.Value(content)))
			case *ast.CodeSpan:
				// 提取行内代码内容
				logger.DebugLogger.Printf("CodeSpan: %s", string(n.Text(content)))
				textSegments = append(textSegments, string(n.Text(content)))
			case *ast.CodeBlock:
				// 提取代码块内容（包括```标记内的代码）
				/*
					lines := n.Lines()
					codeContent := make([]byte, 0)
					for i := 0; i < lines.Len(); i++ {
						seg := lines.At(i)
						logger.DebugLogger.Printf("CodeBlock Line[%d]: %s", i, string(seg.Value(content)))
						codeContent = append(codeContent, seg.Value(content)...)
						codeContent = append(codeContent, '\n') // 保留原始换行
					}
					// 移除末尾多余的换行
					if len(codeContent) > 0 {
						codeContent = codeContent[:len(codeContent)-1]
					}
				*/
				logger.DebugLogger.Printf("FencedCodeBlock Line: %s", string(n.Text(content)))
				codeContent := n.Text(content)
				textSegments = append(textSegments, string(codeContent))
			case *ast.FencedCodeBlock:
				// 提取代码块内容（包括```标记内的代码）
				/*
					lines := n.Lines()
					codeContent := make([]byte, 0)
					for i := 0; i < lines.Len(); i++ {
						seg := lines.At(i)
						logger.DebugLogger.Printf("FencedCodeBlock Line[%d]: %s", i, string(seg.Value(content)))
						codeContent = append(codeContent, seg.Value(content)...)
						codeContent = append(codeContent, '\n') // 保留原始换行
					}
					// 移除末尾多余的换行
					if len(codeContent) > 0 {
						codeContent = codeContent[:len(codeContent)-1]
					}
				*/
				logger.DebugLogger.Printf("FencedCodeBlock Line: %s", string(n.Text(content)))
				codeContent := n.Text(content)
				textSegments = append(textSegments, string(codeContent))
			case *ast.Heading:
				// 提取标题文本（包含所有级别）
				ast.Walk(n, func(child ast.Node, entering bool) (ast.WalkStatus, error) {
					if entering && child.Kind() == ast.KindText {
						logger.DebugLogger.Printf("Heading Text: %s", string(child.(*ast.Text).Value(content)))
						textSegments = append(textSegments, string(child.(*ast.Text).Value(content)))
					}
					return ast.WalkContinue, nil
				})
				return ast.WalkSkipChildren, nil // 跳过子节点避免重复处理
			case *ast.Paragraph:
				// 提取段落文本
				ast.Walk(n, func(child ast.Node, entering bool) (ast.WalkStatus, error) {
					if entering && child.Kind() == ast.KindText {
						logger.DebugLogger.Printf("Paragraph Text: %s", string(child.(*ast.Text).Value(content)))
						textSegments = append(textSegments, string(child.(*ast.Text).Value(content)))
					}
					return ast.WalkContinue, nil
				})
				return ast.WalkSkipChildren, nil // 跳过子节点避免重复处理
			case *ast.ListItem:
				// 提取列表项文本
				ast.Walk(n, func(child ast.Node, entering bool) (ast.WalkStatus, error) {
					if entering && child.Kind() == ast.KindText {
						logger.DebugLogger.Printf("ListItem Text: %s", string(child.(*ast.Text).Value(content)))
						textSegments = append(textSegments, string(child.(*ast.Text).Value(content)))
					}
					return ast.WalkContinue, nil
				})
				return ast.WalkSkipChildren, nil // 跳过子节点避免重复处理
			case *ast.Blockquote:
				// 继续遍历子节点以处理所有内容
				return ast.WalkContinue, nil
			case *ast.List:
				// 处理列表容器，继续遍历子节点
				return ast.WalkContinue, nil
			case *ast.ThematicBreak:
				// 主题分隔线，添加空行分隔
				textSegments = append(textSegments, "")
			case *ast.HTMLBlock:
				// HTML块，跳过处理
				return ast.WalkSkipChildren, nil
			case *ast.Image:
				// 提取图片alt文本
				if n.Lines().Len() > 0 {
					logger.DebugLogger.Printf("Image Alt Text: %s", string(n.Text(content)))
					textSegments = append(textSegments, string(n.Text(content)))
				}
			case *ast.Link:
				// 提取链接文本
				if n.Title != nil {
					textSegments = append(textSegments, string(n.Title))
				}
			case *ast.Emphasis:
				// 处理强调文本节点
				// 内容将通过子Text节点提取
			}
		} else {
			// 块级元素结束时添加换行
			switch node.(type) {
			case *ast.Paragraph, *ast.Heading, *ast.ListItem, *ast.Blockquote, *ast.CodeBlock:
				textSegments = append(textSegments, "\n")
			}
		}
		return ast.WalkContinue, nil
	})

	rawText := strings.Join(textSegments, "\r\n")
	logger.DebugLogger.Printf("Raw Text: %s", rawText)
	return p.processExtractedText(rawText), nil
}

// processExtractedText 处理提取的文本，移除不可见字符并规范化空白
func (p *TextMarkdownParser) processExtractedText(text string) string {
	// 移除不可见字符
	text = invisibleCharsRegex.ReplaceAllString(text, "")
	text = newlineRegex.ReplaceAllString(text, "\n")
	logger.DebugLogger.Printf("1111Raw Text: %s", text)
	// 规范化空白字符
	text = whitespaceRegex.ReplaceAllString(text, " ")
	// 修剪前后空白
	text = strings.TrimSpace(text)
	return text
}

// ParseFile 从Markdown文件中提取纯文本
func (p *TextMarkdownParser) Parse(filePath string) ([]byte, error) {
	content, err := os.ReadFile(filePath)
	if err != nil {
		return []byte{}, fmt.Errorf("无法读取Markdown文件: %w", err)
	}

	data, err := p.ParseMd(content)
	if err != nil {
		return []byte{}, fmt.Errorf("无法解析Markdown文件: %w", err)
	}

	return []byte(data), nil
}

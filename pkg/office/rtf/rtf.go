package rtf

import (
	"fextra/pkg/logger"
	"fmt"
	"io"
	"os"
	"regexp"
	"strings"
)

// TextPosition 表示文本在RTF文件中的位置信息
type TextPosition struct {
	Offset int    // 字节偏移量
	Length int    // 文本长度
	Text   string // 提取的文本内容
}

// OfficeRtfParser RTF文件解析器
type OfficeRtfParser struct{}

// Parse 提取RTF文件中的纯文本内容及位置信息
func (p *OfficeRtfParser) Parse(filename string) ([]byte, error) {
	// 打开文件
	file, err := os.Open(filename)
	if err != nil {
		return []byte{}, fmt.Errorf("无法打开RTF文件: %v", err)
	}
	defer file.Close()

	// 读取文件内容
	content, err := io.ReadAll(file)
	if err != nil {
		return []byte{}, fmt.Errorf("无法读取RTF文件: %v", err)
	}

	// 提取纯文本和位置信息
	extractedText, _ := extractTextWithPositions(string(content))

	return []byte(extractedText), nil
}

// extractTextWithPositions 从RTF内容中提取纯文本及位置信息
func extractTextWithPositions(content string) (string, []TextPosition) {
	var result strings.Builder
	var positions []TextPosition
	var currentText strings.Builder
	var inTextBlock bool
	currentOffset := 0

	// RTF解析状态机
	// 实现思路：采用有限状态机(FSM)模型解析RTF层级结构
	// 通过状态变量跟踪当前解析上下文，区分文本内容与格式控制指令
	state := &parserState{
		inGroup:    false,         // 是否在RTF组内(由{}界定)，初始不在任何组
		groupStack: []groupInfo{}, // 组栈，记录嵌套组信息，支持多层嵌套解析
		inControl:  false,         // 是否在控制字状态(以\开头)，初始不在控制字状态
		controlBuf: "",            // 控制字缓冲区，临时存储当前解析的控制字
		textBuf:    "",            // 文本缓冲区，收集提取的纯文本内容
		Offset:     0,             // 在原始RTF内容中的字节偏移量，用于定位文本位置
	}

	// 逐个字符处理RTF内容
	for i, c := range content {
		state.Offset = i
		state.processChar(c)

		// 记录文本块位置
		if state.inText && !inTextBlock {
			inTextBlock = true
			currentOffset = i
			currentText.Reset()
		} else if !state.inText && inTextBlock {
			inTextBlock = false
			text := currentText.String()
			if strings.TrimSpace(text) != "" {
				positions = append(positions, TextPosition{
					Offset: currentOffset,
					Length: len(text),
					Text:   text,
				})
				logger.Logger.Printf("offset: 0x%x, length: 0x%x, text: %s", currentOffset, len(text), text)
				result.WriteString(text)
			}
		}

		// 收集文本内容
		if state.inText {
			currentText.WriteRune(c)
		}
	}

	// 处理最后一个文本块
	if inTextBlock && currentText.Len() > 0 {
		text := currentText.String()
		if strings.TrimSpace(text) != "" {
			positions = append(positions, TextPosition{
				Offset: currentOffset,
				Length: len(text),
				Text:   text,
			})
			logger.Logger.Printf("111 offset: 0x%x, length: 0x%x, text: %s", currentOffset, len(text), text)
			result.WriteString(text)
		}
		positions = append(positions, TextPosition{
			Offset: currentOffset,
			Length: len(text),
			Text:   text,
		})
		logger.Logger.Printf("111 offset: 0x%x, length: 0x%x, text: %s", currentOffset, len(text), text)
		result.WriteString(text)
	}

	// 应用文本清理规则
	cleanedText := cleanText(result.String())
	return cleanedText, positions
}

// parserState RTF解析器状态
// 状态机核心结构，跟踪RTF文档解析过程中的上下文信息
type parserState struct {
	inGroup    bool        // 是否处于RTF组内(由{}界定)，true表示在组内
	groupStack []groupInfo // 组栈，存储嵌套组信息，支持多层组结构解析
	inControl  bool        // 是否处于控制字解析状态，true表示正在解析\开头的控制字
	controlBuf string      // 控制字缓冲区，临时存储当前解析的控制字内容
	textBuf    string      // 文本缓冲区，收集提取的纯文本内容
	inText     bool        // 是否处于文本内容状态，true表示当前字符为文本内容
	Offset     int         // 在原始RTF内容中的字节偏移量，用于定位文本位置
}

type groupInfo struct {
	// 组在RTF文档中的起始偏移量，用于跟踪解析位置
	startOffset int
	// 定义组类型的控制字（如"fonttbl"表示字体表，"colortbl"表示颜色表）
	typeControl string
	// 标识当前组是否为样式定义组（如字体表、颜色表等），此类组内容不应作为文本提取
	isStyleGroup bool
}

var (
	StyleGroupMap = map[string]struct{}{
		"fonttbl":    {},
		"colortbl":   {},
		"styltbl":    {},
		"listtbl":    {},
		"pict":       {},
		"table":      {},
		"info":       {},
		"revtbl":     {},
		"stylesheet": {},
		"filetable":  {},
		"parfmt":     {},
		"listlevel":  {},
		"pagetbl":    {},
		"shppict":    {},
		"nonshppict": {},
		"field":      {},
		"formfield":  {},
		"listtable":  {},
		"pgptbl":     {},
		"rsidtable":  {},
		"author":     {},
		"operator":   {},
		"themedata":  {}, //非rtf规范，文档主题数据区域，用于存储文档的主题信息。
		"datastore":  {}, //非rtf规范，文档数据存储区域，用于存储文档的元数据、自定义属性等信息。
		"objdata":    {}, //非rtf规范，文档对象数据区域，用于存储文档的对象信息。
		"lsdlocked0": {}, //非rtf规范，Word样式锁定控制字
		"xmlnstbl":   {}, //非rtf规范，XML命名空间表，用于存储文档的XML命名空间信息。
		"ud":         {},
		"upr":        {},
		"xe":         {},
		"urtfN":      {},
		"userprops":  {},
		"vern":       {},
		"version":    {},
		"ts":         {},
		"tsrowd":     {},
		"tscell":     {},
		"wbitmap":    {},
		"wmetafile":  {},
	}
)

func checkStyleGroup(control string) bool {
	_, ok := StyleGroupMap[control]
	return ok
}

// processChar 处理单个字符并更新解析状态
func (s *parserState) processChar(c rune) {
	// 处理组标记
	if c == '{' {
		// 处理组开始前先检查是否有未完成的控制字
		if s.inControl {
			s.processControlWord(s.controlBuf)
			s.inControl = false
		}
		s.inGroup = true
		// 检查是否为样式组（字体表、颜色表等）或继承自父组
		pStyle := len(s.groupStack) > 0 && s.groupStack[len(s.groupStack)-1].isStyleGroup
		isStyleGroup := checkStyleGroup(s.controlBuf) || pStyle
		logger.DebugLogger.Printf("offset: 0x%x, inControl: %v, controlBuf: %s, isStyleGroup: %v, pStyle: %v, groupStack: %d, offset: %d, text: %s",
			s.Offset, s.inControl, s.controlBuf, isStyleGroup, pStyle, len(s.groupStack), len(s.textBuf), s.textBuf)
		// 创建临时组信息，初始标记为非样式组
		tempGroup := groupInfo{startOffset: len(s.textBuf), isStyleGroup: isStyleGroup}
		s.groupStack = append(s.groupStack, tempGroup)

		// 延迟设置样式组状态，等待控制字解析后更新栈顶元素
		// 当控制字解析完成后，调用 updateTopGroupStyle 方法设置正确的样式组状态
		s.inText = false
	} else if c == '}' {
		// 处理组开始前先检查是否有未完成的控制字
		if s.inControl {
			s.processControlWord(s.controlBuf)
			s.inControl = false
		}

		if len(s.groupStack) > 0 {
			// 弹出组信息
			lastGroup := s.groupStack[len(s.groupStack)-1]
			s.groupStack = s.groupStack[:len(s.groupStack)-1]
			s.inGroup = len(s.groupStack) > 0
			// 如果是样式组，清除该组内添加的文本
			if lastGroup.isStyleGroup && lastGroup.startOffset <= len(s.textBuf) {
				s.textBuf = s.textBuf[:lastGroup.startOffset]
			}
		}
		s.inText = false
		// 处理控制字
	} else if c == '\\' {
		// 处理组开始前先检查是否有未完成的控制字
		if s.inControl {
			s.processControlWord(s.controlBuf)
		}
		s.inControl = true
		s.controlBuf = ""
		s.inText = false
	} else if s.inControl {
		// 控制字终止条件：空格、组标记或新控制字开始
		if c == ' ' || c == ';' {
			// 处理控制字并更新控制状态
			s.processControlWord(s.controlBuf)
			s.inControl = false
		} else {
			s.controlBuf += string(c)
		}
	} else {
		// 检查当前是否在样式组中
		inStyleGroup := len(s.groupStack) > 0 && s.groupStack[len(s.groupStack)-1].isStyleGroup
		if !inStyleGroup {
			// 普通文本字符
			s.inText = true
			s.textBuf += string(c)
		} else {
			s.inText = false
		}
	}
}

// processControlWord 处理RTF控制字
// true  -- 在样式组内
// false -- 不在样式组内
func (s *parserState) processControlWord(control string) {
	// 区分文本内容和样式控制字
	// 检查是否为样式组控制字
	isStyleControl := checkStyleGroup(control)
	if isStyleControl && len(s.groupStack) > 0 {
		// 更新栈顶组的样式状态
		lastIdx := len(s.groupStack) - 1
		s.groupStack[lastIdx].isStyleGroup = true
		// 记录样式组类型
		s.groupStack[lastIdx].typeControl = control
		return
	}

	// 样式相关控制字列表
	// 文本结构控制字（保留）
	if strings.HasPrefix(control, "par") || control == "line" || strings.HasPrefix(control, "tab") || strings.HasPrefix(control, "u") {
		// 处理文本结构控制字
		s.textBuf += "\n"
	} else {
		// 忽略样式控制字（字体、颜色、大小等）
		// 可扩展样式控制字列表：f, fs, cf, b, i, u, bold, italic等
		if strings.HasPrefix(control, "f") || strings.HasPrefix(control, "fs") || strings.HasPrefix(control, "cf") ||
			strings.HasPrefix(control, "b") || strings.HasPrefix(control, "i") || strings.HasPrefix(control, "u") {
			// 样式控制字，不添加到文本缓冲区
			return
		}
	}
	return
}

// cleanText 清理提取的文本内容
func cleanText(text string) string {
	// 移除连续重复字符
	//reRepeats := regexp.MustCompile(`(.)\1{3,}`)
	//text = reRepeats.ReplaceAllString(text, "$1")

	// 规范化空白字符
	reWhitespace := regexp.MustCompile(`\s+`)
	text = strings.ReplaceAll(text, "\r\n", "\n")
	text = strings.ReplaceAll(text, "\r", "\n")
	text = reWhitespace.ReplaceAllString(text, " ")
	text = strings.TrimSpace(text)

	return text
}

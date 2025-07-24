package pdf

import (
	"bufio"
	"bytes"
	"fmt"
	"io"
	"os"
	"regexp"
	"strings"

	ledongthucpdf "github.com/ledongthuc/pdf"
	pdfcpu "github.com/pdfcpu/pdfcpu/pkg/api"
	rscpdf "github.com/rsc/pdf"
	"github.com/saintfish/chardet"
	"golang.org/x/text/encoding"
	"golang.org/x/text/encoding/simplifiedchinese"
	"golang.org/x/text/encoding/traditionalchinese"
	"golang.org/x/text/encoding/unicode"
	"golang.org/x/text/transform"

	"fextra/pkg/compressfile"
	"fextra/pkg/logger"
)

// OfficePdfParser PDF文档解析器
type OfficePdfParser struct{}

// Parse 解析PDF文件并提取文本内容
func (p *OfficePdfParser) Parse(filePath string) ([]byte, error) {
	// 尝试ledongthuc/pdf解析
	extractedText, err := p.parseWithStandardLib(filePath)
	if err == nil && len(extractedText) > 0 {
		return extractedText, nil
	}

	// ledongthuc/pdf解析失败，尝试rsc/pdf解析
	logger.Logger.Printf("ledongthuc/pdf解析失败: %v，尝试rsc/pdf解析", err)
	rscText, err := p.parseWithRscPdf(filePath)
	if err == nil && len(rscText) > 0 {
		return rscText, nil
	}

	// rsc/pdf解析失败，尝试pdfcpu解析
	logger.Logger.Printf("rsc/pdf解析失败: %v，尝试pdfcpu解析", err)
	pdfcpuText, err := p.parseWithPdfcpu(filePath)
	if err == nil && len(pdfcpuText) > 0 {
		return pdfcpuText, nil
	}

	// pdfcpu解析失败，尝试二进制解析方案
	logger.Logger.Printf("pdfcpu解析失败: %v，尝试二进制解析", err)
	binaryText, err := p.parseBinaryPDF(filePath)
	if err != nil {
		return []byte{}, fmt.Errorf("所有提取方案均失败: %v", err)
	}

	return binaryText, nil
}

// 使用标准库解析PDF (ledongthuc/pdf)
func (p *OfficePdfParser) parseWithStandardLib(filePath string) ([]byte, error) {
	f, r, err := ledongthucpdf.Open(filePath)
	if err != nil {
		return []byte{}, err
	}
	defer f.Close()

	var textBuilder bytes.Buffer
	pageCount := r.NumPage()

	for i := 1; i <= pageCount; i++ {
		page := r.Page(i)
		if !page.V.IsNull() {
			logger.Logger.Printf("获取第%d页失败", i)
			continue
		}

		content, err := page.GetPlainText(nil)
		if err != nil {
			logger.Logger.Printf("提取第%d页文本失败: %v", i, err)
			continue
		}

		textBuilder.WriteString(content)
		textBuilder.WriteString("\f")
	}

	return textBuilder.Bytes(), nil
}

// 使用rsc/pdf库解析PDF
func (p *OfficePdfParser) parseWithRscPdf(filePath string) ([]byte, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return []byte{}, fmt.Errorf("无法打开文件: %v", err)
	}
	defer file.Close()

	// 解析PDF文件
	pdfReader, err := rscpdf.NewReader(file, 10*1024*1024)
	if err != nil {
		return []byte{}, fmt.Errorf("解析PDF失败: %v", err)
	}

	var textBuilder bytes.Buffer

	// 遍历所有页面
	for pageNum := 1; pageNum <= pdfReader.NumPage(); pageNum++ {
		page := pdfReader.Page(pageNum)
		if page.V.IsNull == nil {
			logger.Logger.Printf("无法获取第%d页", pageNum)
			continue
		}

		// 提取页面文本
		content := page.Content()
		if len(content.Text) == 0 {
			logger.Logger.Printf("第%d页内容为空", pageNum)
			continue
		}

		for _, text := range content.Text {
			textBuilder.WriteString(text.S)
			textBuilder.WriteString("\n")
		}

		textBuilder.WriteString("\f")
	}

	return textBuilder.Bytes(), nil
}

// 使用pdfcpu库解析PDF
func (p *OfficePdfParser) parseWithPdfcpu(filePath string) ([]byte, error) {
	// 创建临时目录
	tmpDir, err := os.MkdirTemp("", "pdf_extract_")
	if err != nil {
		return []byte{}, fmt.Errorf("创建临时目录失败: %v", err)
	}
	defer os.RemoveAll(tmpDir) // 确保程序退出时清理临时目录
	logger.Logger.Printf("临时目录: %s", tmpDir)

	if err = pdfcpu.ExtractContentFile(filePath, tmpDir, nil, nil); err != nil {
		return []byte{}, fmt.Errorf("pdfcpu提取文本失败: %v", err)
	}

	content, cnt, err := compressfile.WalkDir(tmpDir)
	if err != nil {
		return content, err
	}

	logger.Logger.Printf("pdfcpu解析完成，共提取 %d 个页面", cnt)

	return content, nil
}

// 基于二进制解析PDF文本内容
func (p *OfficePdfParser) parseBinaryPDF(filePath string) ([]byte, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return []byte{}, fmt.Errorf("无法打开文件: %v", err)
	}
	defer file.Close()

	// 读取PDF文件头确认格式
	header := make([]byte, 4)
	_, err = file.Read(header)
	if err != nil || !bytes.Equal(header, []byte("%PDF")) {
		return []byte{}, fmt.Errorf("不是有效的PDF文件")
	}

	// 重置文件指针
	_, err = file.Seek(0, io.SeekStart)
	if err != nil {
		return []byte{}, err
	}

	// 使用正则表达式提取文本流内容
	scanner := bufio.NewScanner(file)
	var contentBuffer bytes.Buffer
	textRegex := regexp.MustCompile(`\(([^)]+)\)`)
	streamRegex := regexp.MustCompile(`stream(.*?)endstream`)

	for scanner.Scan() {
		line := scanner.Text()
		// 提取文本对象
		matches := textRegex.FindAllStringSubmatch(line, -1)
		for _, match := range matches {
			if len(match) > 1 {
				contentBuffer.WriteString(match[1])
				contentBuffer.WriteString(" ")
			}
		}

		// 提取流内容
		streamMatches := streamRegex.FindAllStringSubmatch(line, -1)
		for _, match := range streamMatches {
			if len(match) > 1 {
				// 简单处理流中的文本内容
				textContent := textRegex.FindAllStringSubmatch(match[1], -1)
				for _, textMatch := range textContent {
					if len(textMatch) > 1 {
						contentBuffer.WriteString(textMatch[1])
						contentBuffer.WriteString(" ")
					}
				}
			}
		}
	}

	if err := scanner.Err(); err != nil {
		return []byte{}, fmt.Errorf("文件扫描错误: %v", err)
	}

	// 检测并解码文本内容
	extractedText, err := p.detectAndDecodeText(contentBuffer.Bytes())
	if err != nil {
		return []byte{}, err
	}

	// 清理提取的文本
	extractedText = strings.ReplaceAll(extractedText, "\r\n", " ")
	extractedText = strings.ReplaceAll(extractedText, "\n", " ")
	extractedText = regexp.MustCompile(`\s+`).ReplaceAllString(extractedText, " ")

	return []byte(extractedText), nil
}

// detectAndDecodeText 检测文本编码并解码为UTF-8
func (p *OfficePdfParser) detectAndDecodeText(rawData []byte) (string, error) {
	// 检测文本编码
	detector := chardet.NewTextDetector()
	result, err := detector.DetectBest(rawData)
	if err != nil {
		logger.Logger.Printf("编码检测失败: %v，使用默认UTF-8编码", err)
		result = &chardet.Result{Charset: "UTF-8", Confidence: 1.0}
	}

	// 根据检测结果选择解码器
	var decoder encoding.Encoding
	switch strings.ToLower(result.Charset) {
	case "utf-8":
		decoder = encoding.Nop
	case "utf-16le":
		decoder = unicode.UTF16(unicode.LittleEndian, unicode.IgnoreBOM)
	case "utf-16be":
		decoder = unicode.UTF16(unicode.BigEndian, unicode.IgnoreBOM)
	case "gbk", "gb2312", "gb18030":
		decoder = simplifiedchinese.GBK
	case "big5":
		decoder = traditionalchinese.Big5
	default:
		logger.Logger.Printf("不支持的编码格式: %s，使用默认UTF-8解码", result.Charset)
		decoder = encoding.Nop
	}

	// 解码为UTF-8
	decodedBytes, _, err := transform.Bytes(decoder.NewDecoder(), rawData)
	if err != nil {
		return "", fmt.Errorf("文本解码失败: %v", err)
	}

	return string(decodedBytes), nil
}

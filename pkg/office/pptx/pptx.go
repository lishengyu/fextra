package pptx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"

	"fextra/pkg/logger"
)

type OfficePptxParser struct{}

// Parse 提取PPTX文件中的文本内容
func (p *OfficePptxParser) Parse(filename string) ([]byte, error) {
	// 打开ZIP文件
	reader, err := zip.OpenReader(filename)
	if err != nil {
		return []byte{}, fmt.Errorf("无法打开PPTX文件: %v", err)

	}
	defer reader.Close()

	var textBuffer bytes.Buffer

	// 收集所有幻灯片文件
	var slideFiles []*zip.File
	for _, file := range reader.File {
		if filepath.Dir(file.Name) == "ppt/slides" && filepath.Ext(file.Name) == ".xml" {
			// 验证文件名是否符合slide*.xml模式
			if matched, _ := regexp.MatchString(`^slide\d+\.xml$`, filepath.Base(file.Name)); matched {
				slideFiles = append(slideFiles, file)
			} else {
				logger.Logger.Printf("跳过非标准幻灯片文件: %s", file.Name)
			}
		}
	}

	// 按幻灯片编号排序
	sort.Slice(slideFiles, func(i, j int) bool {
		numI := extractSlideNumber(slideFiles[i].Name)
		numJ := extractSlideNumber(slideFiles[j].Name)
		return numI < numJ
	})

	// 处理排序后的幻灯片文件
	for _, file := range slideFiles {
		logger.Logger.Printf("处理幻灯片文件: %v", file.Name)
		// 读取幻灯片内容
		slideContent, err := readZipFile(file)
		if err != nil {
			logger.Logger.Printf("无法读取幻灯片文件 %s: %v", file.Name, err)
			continue
		}

		// 解析幻灯片XML并提取文本
		slideText, err := parseSlideXml(slideContent)
		if err != nil {
			logger.Logger.Printf("无法解析幻灯片XML %s: %v", file.Name, err)
			continue
		}

		// 将幻灯片文本添加到结果中，用分页符分隔
		textBuffer.Write(slideText)
		textBuffer.WriteString("\f") // 使用换页符分隔不同幻灯片
	}

	return textBuffer.Bytes(), nil
}

// extractSlideNumber 从幻灯片文件名中提取编号
func extractSlideNumber(filename string) int {
	re := regexp.MustCompile(`slide(\d+)\.xml`)
	matches := re.FindStringSubmatch(filepath.Base(filename))
	if len(matches) > 1 {
		num, _ := strconv.Atoi(matches[1])
		return num
	}
	return 0 // 无法提取编号时返回0，排在最后
}

// readZipFile 读取ZIP文件中的指定文件内容
func readZipFile(zf *zip.File) ([]byte, error) {
	rc, err := zf.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	content, err := io.ReadAll(rc)
	if err != nil {
		return nil, err
	}

	return content, nil
}

// parseSlideXml 解析幻灯片XML内容并提取文本
func parseSlideXml(xmlContent []byte) ([]byte, error) {
	var slide slideXml
	if err := xml.Unmarshal(xmlContent, &slide); err != nil {
		return []byte{}, err
	}

	ctx, _ := xml.MarshalIndent(slide, "", "  ")
	logger.DebugLogger.Printf("slideXml:\n %s", string(ctx))

	var textBuffer bytes.Buffer

	// 提取所有文本内容
	for _, cSld := range slide.CSld {
		for _, spTree := range cSld.SpTree {
			for _, sp := range spTree.Sp {
				// 仅忽略特定类型的系统占位符
				if sp.Php != nil && sp.Php.Type != nil {
					// 记录占位符类型用于调试
					logger.DebugLogger.Printf("发现占位符类型: %s", *sp.Php.Type)
					// 只跳过系统自动生成的占位符
					if *sp.Php.Type == "sldNum" || *sp.Php.Type == "date" || *sp.Php.Type == "footer" || *sp.Php.Type == "header" {
						logger.DebugLogger.Printf("跳过系统占位符: %s", *sp.Php.Type)
						continue
					}
				}

				for _, txBody := range sp.TxBody {
					for _, p := range txBody.P {
						paraText := extractParagraphText(p)
						if len(paraText) != 0 {
							textBuffer.Write(paraText)
							textBuffer.WriteString("\n")
						}
					}
				}
			}
		}
	}

	return textBuffer.Bytes(), nil
}

// extractParagraphText 提取段落中的文本内容
func extractParagraphText(p para) []byte {
	var paraBuffer bytes.Buffer

	for _, r := range p.R {
		for _, t := range r.T {
			paraBuffer.WriteString(t.Value)
		}
	}

	return paraBuffer.Bytes()
}

// XML结构体定义 - PPTX使用DrawingML命名空间

const drawingMLNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main"
const presentationMLNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main"

// slideXml 幻灯片XML根结构
type slideXml struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/presentationml/2006/main sld"`
	CSld    []cSld   `xml:"http://schemas.openxmlformats.org/presentationml/2006/main cSld"`
}

// cSld 幻灯片内容
type cSld struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/presentationml/2006/main cSld"`
	SpTree  []spTree `xml:"http://schemas.openxmlformats.org/presentationml/2006/main spTree"`
}

// spTree 形状树
type spTree struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/presentationml/2006/main spTree"`
	Sp      []sp     `xml:"sp"`
}

// sp 形状
type sp struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/presentationml/2006/main sp"`
	Php     *php     `xml:"http://schemas.openxmlformats.org/presentationml/2006/main ph"` // 占位符标识
	TxBody  []txBody `xml:"http://schemas.openxmlformats.org/presentationml/2006/main txBody"`
}

// php 占位符属性
type php struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/presentationml/2006/main ph"`
	Type    *string  `xml:"type,attr"`
}

// txBody 文本主体
type txBody struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/presentationml/2006/main txBody"`
	P       []para   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main p"` // 段落
}

// para 段落
type para struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/main p"`
	R       []r      `xml:"http://schemas.openxmlformats.org/drawingml/2006/main r"` // 文本 run
}

// r 文本 run
type r struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/main r"`
	T       []t      `xml:"http://schemas.openxmlformats.org/drawingml/2006/main t"` // 文本内容
}

// t 文本元素
type t struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/main t"`
	Value   string   `xml:",chardata"`
}

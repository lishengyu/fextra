package xlsx

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

// OfficeXlsxParser XLSX文件解析器
type OfficeXlsxParser struct{}

// Parse 提取XLSX文件中的文本内容
func (p *OfficeXlsxParser) Parse(filename string) ([]byte, error) {
	// 打开ZIP文件
	reader, err := zip.OpenReader(filename)
	if err != nil {
		return []byte{}, fmt.Errorf("无法打开XLSX文件: %v", err)
	}
	defer reader.Close()

	// 读取共享字符串表
	sharedStrings, err := readSharedStrings(reader)
	if err != nil {
		// 非致命错误，继续处理
		logger.Logger.Printf("读取共享字符串表失败: %v", err)
	}

	// 收集所有工作表文件
	var sheetFiles []*zip.File
	for _, file := range reader.File {
		if filepath.Dir(file.Name) == "xl/worksheets" && filepath.Ext(file.Name) == ".xml" {
			// 验证文件名是否符合sheet*.xml模式
			if matched, _ := regexp.MatchString(`^sheet\d+\.xml$`, filepath.Base(file.Name)); matched {
				sheetFiles = append(sheetFiles, file)
			} else {
				logger.Logger.Printf("跳过非标准工作表文件: %s", file.Name)
			}
		}
	}

	// 按工作表编号排序
	sort.Slice(sheetFiles, func(i, j int) bool {
		numI := extractSheetNumber(sheetFiles[i].Name)
		numJ := extractSheetNumber(sheetFiles[j].Name)
		return numI < numJ
	})

	var textBuffer bytes.Buffer

	// 处理排序后的工作表文件
	for _, file := range sheetFiles {
		logger.Logger.Printf("处理工作表文件: %v", file.Name)
		// 读取工作表内容
		sheetContent, err := readZipFile(file)
		if err != nil {
			logger.Logger.Printf("无法读取工作表文件 %s: %v", file.Name, err)
			continue
		}

		// 解析工作表XML并提取文本
		sheetText, err := parseSheetXml(sheetContent, sharedStrings)
		if err != nil {
			logger.Logger.Printf("无法解析工作表XML %s: %v", file.Name, err)
			continue
		}

		// 将工作表文本添加到结果中，用分页符分隔
		textBuffer.WriteString(fmt.Sprintf("=== 工作表: %s ===\n", filepath.Base(file.Name)))
		textBuffer.Write(sheetText)
		textBuffer.WriteString("\n\f\n") // 使用换页符分隔不同工作表
	}

	return textBuffer.Bytes(), nil
}

// readSharedStrings 读取共享字符串表
func readSharedStrings(reader *zip.ReadCloser) ([]string, error) {
	for _, file := range reader.File {
		if file.Name == "xl/sharedStrings.xml" {
			content, err := readZipFile(file)
			if err != nil {
				return nil, err
			}

			var sst sharedStrings
			if err := xml.Unmarshal(content, &sst); err != nil {
				return nil, err
			}

			// 提取共享字符串
			strings := make([]string, len(sst.Si))
			for i, si := range sst.Si {
				strings[i] = si.T.Value
			}
			return strings, nil
		}
	}
	return []string{}, nil // 没有共享字符串表
}

// parseSheetXml 解析工作表XML并提取文本
func parseSheetXml(xmlContent []byte, sharedStrings []string) ([]byte, error) {
	var worksheet worksheet
	if err := xml.Unmarshal(xmlContent, &worksheet); err != nil {
		return []byte{}, err
	}

	var sheetBuffer bytes.Buffer

	// 遍历所有行
	for _, row := range worksheet.SheetData.Row {
		var rowBuffer bytes.Buffer
		// 遍历行中的单元格
		for _, c := range row.C {
			// 获取单元格值
			cellValue := getCellValue(c, sharedStrings)
			if cellValue != "" {
				if rowBuffer.Len() > 0 {
					rowBuffer.WriteString("\t") // 使用制表符分隔单元格
				}
				rowBuffer.WriteString(cellValue)
			}
		}
		// 添加行文本（如果不为空）
		if rowBuffer.Len() > 0 {
			sheetBuffer.Write(rowBuffer.Bytes())
			sheetBuffer.WriteString("\n") // 使用换行符分隔行
		}
	}

	return sheetBuffer.Bytes(), nil
}

// getCellValue 获取单元格值，处理共享字符串引用
func getCellValue(c cell, sharedStrings []string) string {
	if c.T == "s" && c.V != "" {
		// 共享字符串引用
		index, err := strconv.Atoi(c.V)
		if err == nil && index >= 0 && index < len(sharedStrings) {
			return sharedStrings[index]
		}
	}
	// 直接返回单元格值或其他类型数据
	return c.V
}

// extractSheetNumber 从工作表文件名中提取编号
func extractSheetNumber(filename string) int {
	re := regexp.MustCompile(`sheet(\d+)\.xml`)
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

// XML结构体定义 - XLSX使用SpreadsheetML命名空间

const spreadsheetMLNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

// worksheet 工作表XML根结构
type worksheet struct {
	XMLName   xml.Name  `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	SheetData sheetData `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main sheetData"`
}

// sheetData 工作表数据
type sheetData struct {
	Row []row `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main row"`
}

// row 行
type row struct {
	C []cell `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main c"` // 单元格
}

// cell 单元格
type cell struct {
	V string `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main v"` // 单元格值
	T string `xml:"t,attr"`                                                      // 单元格类型 (s表示共享字符串)
}

// sharedStrings 共享字符串表
type sharedStrings struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main sst"`
	Si      []si     `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main si"`
}

// si 共享字符串项
type si struct {
	T t `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main t"`
}

// t 文本元素
type t struct {
	Value string `xml:",chardata"`
}

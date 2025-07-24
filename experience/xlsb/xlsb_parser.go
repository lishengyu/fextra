package xlsb

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"fextra/pkg/logger"
	"fmt"
	"io"
	"strconv"
	"strings"

	"golang.org/x/text/encoding/unicode"
)

// XLSB记录类型常量定义
const (
	BRT_CellBlank uint32 = 1  // 空白单元格
	BRT_CellRk    uint32 = 2  // 数值型单元格
	BRT_CellBool  uint32 = 4  // 布尔型单元格
	BRT_CellIstr  uint32 = 6  // 内联字符串单元格
	BRT_CellIsst  uint32 = 7  // 共享字符串单元格
	BRT_SstItem   uint32 = 19 // 共享字符串项

	BRT_CellFormula uint32 = 3 // 公式单元格
)

// SharedStringTable 共享字符串表
type SharedStringTable struct {
	items []string
}

// OfficeXlsbParser XLSB解析器
type OfficeXlsbParser struct {
	sharedStrings *SharedStringTable
}

// Parse 解析XLSB文件并提取文本内容
func (p *OfficeXlsbParser) Parse(filePath string) ([]byte, error) {
	// 初始化共享字符串表
	p.sharedStrings = &SharedStringTable{items: make([]string, 0)}

	// 打开ZIP格式的XLSB文件
	zipReader, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, fmt.Errorf("打开XLSB文件失败: %w", err)
	}
	defer zipReader.Close()

	var textBuilder bytes.Buffer

	// 1. 解析共享字符串表
	if err := p.parseSharedStrings(zipReader); err != nil {
		logger.Logger.Printf("解析共享字符串表失败: %v", err)
		// 非致命错误，继续解析工作表
	}

	// 2. 解析所有工作表
	for _, file := range zipReader.File {
		if strings.HasPrefix(file.Name, "xl/worksheets/") && strings.HasSuffix(file.Name, ".bin") {
			if err := p.parseWorksheet(file, &textBuilder); err != nil {
				logger.Logger.Printf("解析工作表 %s 失败: %v", file.Name, err)
			}
		}
	}

	return textBuilder.Bytes(), nil
}

// parseSharedStrings 解析共享字符串表
func (p *OfficeXlsbParser) parseSharedStrings(zipReader *zip.ReadCloser) error {
	for _, file := range zipReader.File {
		logger.Logger.Printf("工作表名: %s", file.Name)
		if file.Name == "xl/sharedStrings.bin" {
			f, err := file.Open()
			if err != nil {
				return err
			}
			defer f.Close()

			// 解析共享字符串二进制内容
			return p.parseSstBinary(f)
		}
	}
	return nil // 没有共享字符串表
}

// parseWorksheet 解析单个工作表
func (p *OfficeXlsbParser) parseWorksheet(file *zip.File, textBuilder *bytes.Buffer) error {
	f, err := file.Open()
	if err != nil {
		return err
	}
	defer f.Close()

	// 写入工作表名称
	sheetName := strings.TrimSuffix(strings.TrimPrefix(file.Name, "xl/worksheets/"), ".bin")
	textBuilder.WriteString("工作表: " + sheetName + "\n")

	// 解析工作表二进制内容
	return p.parseWorksheetBinary(f, textBuilder)
}

// parseSstBinary 解析共享字符串二进制数据
func (p *OfficeXlsbParser) parseSstBinary(reader io.Reader) error {
	// XLSB记录头: 4字节类型 + 4字节大小
	var recordHeader [8]byte

	for {
		// 读取记录头
		n, err := reader.Read(recordHeader[:])
		if err != nil {
			if err == io.EOF {
				break
			}
			return fmt.Errorf("读取SST记录头失败: %w", err)
		}
		if n < 8 {
			return fmt.Errorf("SST记录头不完整")
		}

		// 解析记录类型和大小
		recordType := binary.LittleEndian.Uint32(recordHeader[:4])
		recordSize := binary.LittleEndian.Uint32(recordHeader[4:8])

		// 读取记录数据
		recordData := make([]byte, recordSize)
		if _, err := io.ReadFull(reader, recordData); err != nil {
			return fmt.Errorf("读取SST记录数据失败: %w", err)
		}

		// 处理共享字符串项记录
		if recordType == BRT_SstItem {
			str, err := parseXLUnicodeString(recordData)
			if err != nil {
				logger.Logger.Printf("解析共享字符串失败: %v", err)
				continue
			}
			p.sharedStrings.items = append(p.sharedStrings.items, str)
		}
	}

	logger.Logger.Printf("已解析共享字符串: %d 项", len(p.sharedStrings.items))
	return nil
}

// parseWorksheetBinary 解析工作表二进制数据
func (p *OfficeXlsbParser) parseWorksheetBinary(reader io.Reader, textBuilder *bytes.Buffer) error {
	// XLSB记录头: 4字节类型 + 4字节大小
	var recordHeader [8]byte
	var currentRow uint32
	var currentRowCells []string

	for {
		// 读取记录头
		n, err := reader.Read(recordHeader[:])
		if err != nil {
			if err == io.EOF {
				break
			}
			return fmt.Errorf("读取工作表记录头失败: %w", err)
		}
		if n < 8 {
			return fmt.Errorf("工作表记录头不完整")
		}

		// 解析记录类型和大小
		recordType := binary.LittleEndian.Uint32(recordHeader[:4])
		recordSize := binary.LittleEndian.Uint32(recordHeader[4:8])

		// 读取记录数据
		recordData := make([]byte, recordSize)
		if _, err := io.ReadFull(reader, recordData); err != nil {
			return fmt.Errorf("读取工作表记录数据失败: %w", err)
		}

		logger.Logger.Printf("记录类型: %d, 记录大小: %d", recordType, recordSize)
		// 处理不同类型的记录
		switch recordType {
		case BRT_CellRk:
			if err := p.handleCellRk(recordData, &currentRow, &currentRowCells, textBuilder); err != nil {
				logger.Logger.Printf("处理RK单元格失败: %v", err)
				continue
			}
		case BRT_CellBool:
			if err := p.handleCellBool(recordData, &currentRow, &currentRowCells, textBuilder); err != nil {
				logger.Logger.Printf("处理布尔单元格失败: %v", err)
				continue
			}

		case BRT_CellIstr:
			if err := p.handleCellIstr(recordData, &currentRow, &currentRowCells, textBuilder); err != nil {
				logger.Logger.Printf("处理内联字符串单元格失败: %v", err)
				continue
			}
		case BRT_CellFormula:
			/*
				// 公式单元格(简化处理)
				if len(recordData) < 16 {
					logger.Logger.Printf("公式单元格数据不完整")
					continue
				}
				row := binary.LittleEndian.Uint32(recordData[0:4])
				col := binary.LittleEndian.Uint32(recordData[4:8])
				// 提取公式缓存值
				// 公式记录结构: 行(4) + 列(4) + 选项(4) + 公式长度(4) + 公式数据 + 缓存值
				if len(recordData) > 20 {
					// 简单判断是否包含缓存值(实际需根据选项判断)
					cacheType := recordData[16]
					switch cacheType {
					case 0x00:
						// 无缓存值
						value = "[公式]"
					case 0x01:
						// 数值缓存
						cacheValue := binary.LittleEndian.Float64(recordData[20:28])
						value = strconv.FormatFloat(cacheValue, 'f', -1, 64)
					case 0x02:
						// 字符串缓存
						strData := recordData[20:]
						value, _ = parseXLUnicodeString(strData)
					default:
						value = "[公式]"
					}
				} else {
					value = "[公式]"
				}
				// 处理行切换
				if row != currentRow && len(currentRowCells) > 0 {
					textBuilder.WriteString(strings.Join(currentRowCells, "\t") + "\n")
					currentRowCells = make([]string, 0)
				}
				currentRow = row
				// 处理列索引
				if int(col) >= len(currentRowCells) {
					need := int(col) - len(currentRowCells) + 1
					currentRowCells = append(currentRowCells, make([]string, need)...)
				}
				currentRowCells[col] = value
			*/
		case BRT_CellIsst:
			// 共享字符串单元格
			if len(recordData) < 12 {
				logger.Logger.Printf("单元格记录数据不完整")
				continue
			}

			// 解析行号和列号 (前8字节)
			row := binary.LittleEndian.Uint32(recordData[0:4])
			col := binary.LittleEndian.Uint32(recordData[4:8])
			isst := binary.LittleEndian.Uint32(recordData[8:12])

			// 切换行时输出之前的行数据
			if row != currentRow && len(currentRowCells) > 0 {
				textBuilder.WriteString(strings.Join(currentRowCells, "\t") + "\n")
				currentRowCells = make([]string, 0)
			}
			currentRow = row

			var value string
			// 获取共享字符串
			// 通过isst索引获取共享字符串
			if int(isst) < len(p.sharedStrings.items) {
				value = p.sharedStrings.items[isst]
			} else {
				logger.Logger.Printf("共享字符串索引越界: %d", isst)
				value = ""
			}
			// 将单元格值添加到行数据
			currentRowCells = append(currentRowCells, value)

			if len(currentRowCells) > 0 {
				textBuilder.WriteString(strings.Join(currentRowCells, "\t") + "\n")
			}
			for int(col) >= len(currentRowCells) {
				currentRowCells = append(currentRowCells, "")
			}
			currentRowCells[col] = value

		case BRT_CellBlank:
			// 空白单元格，暂不处理
			continue
		}
	}

	// 输出最后一行数据
	if len(currentRowCells) > 0 {
		textBuilder.WriteString(strings.Join(currentRowCells, "\t") + "\n")
	}

	return nil
}

// parseXLUnicodeString 解析XLUnicodeString结构
func parseXLUnicodeString(data []byte) (string, error) {
	if len(data) < 3 {
		return "", fmt.Errorf("字符串数据不完整")
	}

	// 解析字符串长度和编码标志
	cch := binary.LittleEndian.Uint16(data[0:2])
	highByte := (data[2] >> 7) & 0x01
	rgbStart := 3

	// 计算字符串字节长度
	var strLen int
	if highByte == 1 {
		strLen = int(cch) * 2
	} else {
		strLen = int(cch)
	}

	// 验证缓冲区长度
	if rgbStart+strLen > len(data) {
		return "", fmt.Errorf("字符串长度不足, 需要%d字节, 实际%d字节", rgbStart+strLen, len(data))
	}

	// 提取字符串内容
	if highByte == 1 {
		// 双字节Unicode (UTF-16LE)
		utf16Bytes := data[rgbStart : rgbStart+strLen]
		decoder := unicode.UTF16(unicode.LittleEndian, unicode.UseBOM).NewDecoder()
		return decoder.String(string(utf16Bytes))
	} else {
		// 单字节字符
		return string(data[rgbStart : rgbStart+strLen]), nil
	}
}

// 数值型单元格处理函数
func (p *OfficeXlsbParser) handleCellRk(recordData []byte, currentRow *uint32, currentRowCells *[]string, textBuilder *bytes.Buffer) error {
	if len(recordData) < 12 {
		logger.Logger.Printf("RK单元格记录数据不完整")
		return nil
	}
	row := binary.LittleEndian.Uint32(recordData[0:4])
	rkValue := binary.LittleEndian.Uint32(recordData[8:12])
	value := strconv.FormatFloat(float64(rkValue)/65536.0, 'f', -1, 64)
	if row != *currentRow && len(*currentRowCells) > 0 {
		textBuilder.WriteString(strings.Join(*currentRowCells, "\t") + "\n")
		*currentRowCells = make([]string, 0)
	}
	*currentRow = row
	*currentRowCells = append(*currentRowCells, value)
	return nil
}

// 布尔型单元格处理函数
func (p *OfficeXlsbParser) handleCellBool(recordData []byte, currentRow *uint32, currentRowCells *[]string, textBuilder *bytes.Buffer) error {
	if len(recordData) < 9 {
		logger.Logger.Printf("布尔单元格记录数据不完整")
		return nil
	}
	row := binary.LittleEndian.Uint32(recordData[0:4])
	boolValue := recordData[8] != 0
	value := strconv.FormatBool(boolValue)
	if row != *currentRow && len(*currentRowCells) > 0 {
		textBuilder.WriteString(strings.Join(*currentRowCells, "\t") + "\n")
		*currentRowCells = make([]string, 0)
	}
	*currentRow = row
	*currentRowCells = append(*currentRowCells, value)
	return nil
}

// 内联字符串单元格处理函数
func (p *OfficeXlsbParser) handleCellIstr(recordData []byte, currentRow *uint32, currentRowCells *[]string, textBuilder *bytes.Buffer) error {
	if len(recordData) < 9 {
		logger.Logger.Printf("内联字符串单元格数据不完整")
		return nil
	}
	row := binary.LittleEndian.Uint32(recordData[0:4])
	col := binary.LittleEndian.Uint32(recordData[4:8])
	strData := recordData[8:]
	value, err := parseXLUnicodeString(strData)
	if err != nil {
		logger.Logger.Printf("解析内联字符串失败: %v", err)
		value = ""
	}
	if row != *currentRow && len(*currentRowCells) > 0 {
		textBuilder.WriteString(strings.Join(*currentRowCells, "\t") + "\n")
		*currentRowCells = make([]string, 0)
	}
	*currentRow = row
	if int(col) >= len(*currentRowCells) {
		need := int(col) - len(*currentRowCells) + 1
		*currentRowCells = append(*currentRowCells, make([]string, need)...)
	}
	(*currentRowCells)[col] = value
	return nil
}

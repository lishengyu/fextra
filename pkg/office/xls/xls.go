package xls

import (
	"bytes"
	"fmt"

	exls "github.com/extrame/xls"
)

type OfficeXlsParser struct{}

func (p *OfficeXlsParser) Parse(filePath string) ([]byte, error) {
	content, err := ExtractTextFromXLS(filePath)
	if err != nil {
		return nil, err
	}

	return []byte(content), nil
}

func ExtractTextFromXLS(filePath string) ([]byte, error) {
	// 打开文件并指定编码
	file, err := exls.Open(filePath, "utf-8")
	if err != nil {
		return []byte{}, fmt.Errorf("文件打开失败: %v", err)
	}

	var content bytes.Buffer

	// 遍历所有工作表
	for sheetIndex := 0; sheetIndex < file.NumSheets(); sheetIndex++ {
		sheet := file.GetSheet(sheetIndex)
		if sheet == nil {
			continue // 跳过空工作表
		}

		// 添加工作表标题
		content.WriteString(fmt.Sprintf("\n--- 工作表 %d: %s ---\n", sheetIndex+1, sheet.Name))

		// 遍历行 (MaxRow+1 兼容空行)
		for rowIndex := 0; rowIndex <= int(sheet.MaxRow); rowIndex++ {
			row := sheet.Row(rowIndex)
			if row == nil {
				continue // 跳过空行
			}

			// 构建当前行文本
			var rowText bytes.Buffer
			for colIndex := 0; colIndex < row.LastCol(); colIndex++ {
				cell := row.Col(colIndex)
				if cell != "" { // 跳过空单元格
					rowText.WriteString(cell)
					if colIndex < row.LastCol()-1 {
						rowText.WriteString("\t") // 单元格分隔符
					}
				}
			}

			// 添加非空行内容
			if rowText.Len() > 0 {
				content.Write(rowText.Bytes())
				content.WriteString("\n")
			}
		}
	}

	return content.Bytes(), nil
}

package vsd

import (
	"bytes"
	"fmt"
	"os"

	// 复用项目已有Office解析库
	"fextra/pkg/logger"
	// 文档转换工具
)

type OfficeVsdParser struct{}

func (p *OfficeVsdParser) Parse(filePath string) ([]byte, error) {
	content, err := StdLibExtractText(filePath)
	if err == nil && content != "" {
		return []byte(content), nil
	}

	logger.Logger.Printf("标准库解析VSD文件失败: %v", err)

	content, err = BinaryExtractText(filePath)
	if err == nil && content != "" {
		return []byte(content), nil
	}

	logger.Logger.Printf("二进制解析VSD文件失败: %v", err)

	return []byte{}, err
}

// StdLibExtractText 从VSD文件中提取文本内容
func StdLibExtractText(filePath string) (string, error) {
	// 使用docconv库解析VSD文件
	// 内部会调用libreoffice或catdoc等外部工具
	/*
		text, meta, err := docconv.Convert(filePath)
		if err != nil {
			logger.Logger.Printf("VSD文件解析失败: %v", err)
			// 尝试备选方案: 直接读取二进制内容中的文本片段
			return BinaryExtractText(filePath), err
		}

		// 合并元数据和正文内容
		fullText := fmt.Sprintf("Title: %s\nSubject: %s\nCreator: %s\n\n%s",
			meta.Title, meta.Subject, meta.Creator, text)
	*/
	return "", nil
}

// BinaryExtractText 二进制文件文本提取备选方案
func BinaryExtractText(filePath string) (string, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return "", fmt.Errorf("无法打开文件: %v", err)
	}
	defer file.Close()

	// 读取文件前1MB内容用于文本提取
	buf := make([]byte, 1024*1024)
	n, _ := file.Read(buf)
	content := buf[:n]

	// 提取可打印字符
	var textBuilder bytes.Buffer
	for _, b := range content {
		if b >= 32 && b <= 126 || b == 10 || b == 13 {
			textBuilder.WriteByte(b)
		}
	}

	return textBuilder.String(), nil
}

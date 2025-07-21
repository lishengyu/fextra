package vsdx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"fextra/pkg/logger"
)

var (
	// 常见图片文件扩展名，后续可再补充
	imageExtensions = map[string]bool{
		".png":  true,
		".jpg":  true,
		".jpeg": true,
		".tif":  true,
		".tiff": true,
		".webp": true,
		".wbmp": true,
		".fpx":  true,
		".pbm":  true,
		".pgm":  true,
		".gif":  true,
		".bmp":  true,
		".svg":  true,
	}
)

type OfficeVsdxParser struct{}

// 用于提取VSDX文件中的文本内容
func (v *OfficeVsdxParser) Parse(filePath string) ([]byte, error) {
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return []byte{}, fmt.Errorf("无法打开VSDX文件: %v", err)
	}
	defer reader.Close()

	var textBuilder bytes.Buffer

	// 遍历ZIP中的所有文件
	for _, file := range reader.File {
		// 只处理页面内容文件
		if strings.HasPrefix(file.Name, "visio/pages/") && strings.HasSuffix(file.Name, ".xml") {
			text, err := extractTextFromPageXML(file)
			if err != nil {
				logger.Logger.Printf("处理页面文件失败 %s: %v", file.Name, err)
				continue
			}
			textBuilder.Write(text)
			textBuilder.WriteString("\n")
		}
	}

	images, err := ExtractImages(filePath)
	if err != nil {
		logger.Logger.Printf("检测图片文件失败 %s: %v", filePath, err)
	}
	logger.Logger.Printf("文件 %s 包含图片: %d", filePath, images)

	return textBuilder.Bytes(), nil
}

// 从页面XML中提取文本
func extractTextFromPageXML(file *zip.File) ([]byte, error) {
	fileReader, err := file.Open()
	if err != nil {
		return []byte{}, err
	}
	defer fileReader.Close()

	decoder := xml.NewDecoder(fileReader)
	decoder.Strict = false // 忽略XML命名空间和格式问题

	var textBuilder bytes.Buffer
	var inTextElement bool

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return textBuilder.Bytes(), err
		}

		switch t := token.(type) {
		case xml.StartElement:
			// 检测文本元素（处理命名空间）
			if strings.HasSuffix(t.Name.Local, "Text") {
				inTextElement = true
			}
		case xml.EndElement:
			if strings.HasSuffix(t.Name.Local, "Text") {
				inTextElement = false
			}
		case xml.CharData:
			if inTextElement {
				textBuilder.Write(t)
			}
		}
	}

	return textBuilder.Bytes(), nil
}

// HasImages 检查VSDX文件中是否包含图片
func HasImages(filePath string) (bool, error) {
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return false, fmt.Errorf("无法打开VSDX文件: %v", err)
	}
	defer reader.Close()

	// 检查media目录中的文件
	for _, file := range reader.File {
		if strings.HasPrefix(file.Name, "visio/media/") {
			ext := strings.ToLower(filepath.Ext(file.Name))
			if imageExtensions[ext] {
				return true, nil
			}
		}
	}

	return false, nil
}

// ExtractImages 从VSDX文件中提取所有图片并保存到指定目录
func ExtractImages(filePath string) (int, error) {
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return 0, fmt.Errorf("无法打开VSDX文件: %v", err)
	}
	defer reader.Close()

	// 创建临时目录
	tmpDir, err := os.MkdirTemp("", "vsdx_extract_")
	if err != nil {
		return 0, fmt.Errorf("创建临时目录失败: %v", err)
	}

	// todo: fixme later  后期确认提取文件如何处理
	//defer os.RemoveAll(tmpDir) // 确保程序退出时清理临时目录
	logger.Logger.Printf("临时目录: %s", tmpDir)

	filesCnt := 0
	// 提取media目录中的图片文件
	for _, file := range reader.File {
		if strings.HasPrefix(file.Name, "visio/media/") {
			ext := strings.ToLower(filepath.Ext(file.Name))
			if imageExtensions[ext] {
				filesCnt++

				// 打开ZIP中的文件
				zipFile, err := file.Open()
				if err != nil {
					logger.Logger.Printf("无法打开图片文件 %s: %v", file.Name, err)
					continue
				}
				defer zipFile.Close()

				// 创建输出文件
				fileName := filepath.Base(file.Name)
				outputPath := filepath.Join(tmpDir, fileName) // 可以优化，防止路径注入
				outFile, err := os.Create(outputPath)
				if err != nil {
					logger.Logger.Printf("无法创建输出文件 %s: %v", outputPath, err)
					continue
				}
				defer outFile.Close()

				// 复制文件内容
				if _, err := io.Copy(outFile, zipFile); err != nil {
					logger.Logger.Printf("无法保存图片文件 %s: %v", outputPath, err)
					continue
				}

				logger.Logger.Printf("成功提取图片: %s", outputPath)
			}
		}
	}

	return filesCnt, nil
}

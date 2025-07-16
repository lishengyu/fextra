package compressfile

import (
	"compress/gzip"
	"fextra/internal"
	"fextra/pkg/logger"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
)

type GzFileParser struct{}

func (p *GzFileParser) Parse(filePath string) ([]byte, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return []byte{}, fmt.Errorf("无法打开文件: %v", err)
	}
	defer file.Close()

	return parseGzFromReader(file, filePath)
}

func init() {
	// GZ相关类型: 19(gz), 20(tar.gz)
	internal.RegisterParser(internal.FileTypeTARGZ, &GzFileParser{})
	internal.RegisterParser(internal.FileTypeGZ, &GzFileParser{})
}

func writeGzFile(gz *gzip.Reader, path string) error {
	// 创建父目录（如果不存在）
	parentDir := filepath.Dir(path)
	if err := os.MkdirAll(parentDir, 0755); err != nil {
		return err
	}

	// 创建文件
	file, err := os.OpenFile(path, os.O_CREATE|os.O_WRONLY|os.O_TRUNC, os.ModePerm)
	if err != nil {
		return err
	}
	defer file.Close()

	// 流式复制内容（避免内存溢出）
	if _, err := io.Copy(file, gz); err != nil {
		return err
	}
	return nil
}

// parseGzFromReader 从io.Reader解析gz内容并返回格式化字符串
func parseGzFromReader(reader io.Reader, filename string) ([]byte, error) {
	gzReader, err := gzip.NewReader(reader)
	if err != nil {
		return []byte{}, fmt.Errorf("创建gzip reader失败: %v", err)
	}
	defer gzReader.Close()

	// 创建临时目录
	tmpDir, err := os.MkdirTemp("", "gz_extract_")
	if err != nil {
		return []byte{}, fmt.Errorf("创建临时目录失败: %v", err)
	}
	defer os.RemoveAll(tmpDir) // 确保程序退出时清理临时目录
	logger.Logger.Printf("临时目录: %s", tmpDir)

	original := gzReader.Header.Name
	if original == "" {
		if strings.HasSuffix(filename, ".tar.gz") {
			original = "default_gz_file_name.tar"
		} else {
			original = "default_gz_file_name.txt"
		}
	}
	logger.Logger.Printf("原始文件名: %s", original)

	safePath := filepath.Join(tmpDir, sanitizePath(original))

	if err = writeGzFile(gzReader, safePath); err != nil {
		return []byte{}, err
	}

	content, files, err := walkDir(tmpDir)
	if err != nil {
		return content, err
	}
	logger.Logger.Printf("gz文件解析完成，共提取 %d 个文件(一级目录)", files)
	return content, nil
}

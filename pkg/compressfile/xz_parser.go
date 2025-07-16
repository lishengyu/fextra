package compressfile

import (
	"bytes"
	"fextra/internal"
	"fextra/pkg/logger"
	"fmt"
	"io"
	"os"
	"path/filepath"

	"github.com/ulikunitz/xz"
)

type XzFileParser struct{}

func (p *XzFileParser) Parse(filePath string) ([]byte, error) {
	var content bytes.Buffer

	file, err := os.Open(filePath)
	if err != nil {
		return content.Bytes(), fmt.Errorf("无法打开文件: %v", err)
	}
	defer file.Close()

	return parseXzFromReader(file, filePath)
}

func init() {
	// XZ相关类型:29(xz)
	internal.RegisterParser(internal.FileTypeXZ, &XzFileParser{})
}

func WriteXzFile(reader *xz.Reader, path string, mode os.FileMode) error {
	file, err := os.OpenFile(path, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, mode)
	if err != nil {
		return err
	}
	defer file.Close()

	_, err = io.Copy(file, reader)
	return err
}

func parseXzFromReader(reader io.Reader, filename string) ([]byte, error) {
	xzReader, err := xz.NewReader(reader)
	if err != nil {
		return []byte{}, err
	}

	// 创建临时目录
	tmpDir, err := os.MkdirTemp("", "xz_extract_")
	if err != nil {
		return []byte{}, fmt.Errorf("创建临时目录失败: %v", err)
	}
	defer os.RemoveAll(tmpDir) // 确保程序退出时清理临时目录
	logger.Logger.Printf("临时目录: %s", tmpDir)

	original := filepath.Base(filename[:len(filename)-len(".xz")])
	safePath := filepath.Join(tmpDir, sanitizePath(original))

	if err = WriteXzFile(xzReader, safePath, os.ModePerm); err != nil {
		return []byte{}, err
	}

	content, cnt, err := walkDir(tmpDir)
	if err != nil {
		return content, err
	}

	logger.Logger.Printf("xz文件解析完成，共提取 %d 个文件(一级目录)", cnt)
	return content, nil
}

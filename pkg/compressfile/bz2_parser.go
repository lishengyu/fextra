package compressfile

import (
	"bytes"
	"fextra/internal"
	"fextra/pkg/logger"
	"fmt"
	"io"
	"io/fs"
	"os"
	"path/filepath"

	"compress/bzip2"
)

type Bz2FileParser struct{}

func (p *Bz2FileParser) Parse(filePath string) ([]byte, error) {
	var content bytes.Buffer

	file, err := os.Open(filePath)
	if err != nil {
		return content.Bytes(), fmt.Errorf("无法打开文件: %v", err)
	}
	defer file.Close()

	return parseBz2FromReader(file, filePath)
}

func init() {
	// BZ2相关类型: 24(bz2)
	internal.RegisterParser(internal.FileTypeBZ2, &Bz2FileParser{})
}

func WriteBz2File(rc io.Reader, safePath string, mode fs.FileMode) error {
	// 创建目标文件
	dstFile, err := os.OpenFile(safePath, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, mode)
	if err != nil {
		return fmt.Errorf("创建文件 %s 失败: %v", safePath, err)
	}
	defer dstFile.Close()

	// 复制文件内容
	if _, err := io.Copy(dstFile, rc); err != nil {
		return fmt.Errorf("复制文件 %s 内容失败: %v", safePath, err)
	}
	return nil
}

func parseBz2FromReader(reader io.Reader, filename string) ([]byte, error) {
	bz2Reader := bzip2.NewReader(reader)

	// 创建临时目录
	tmpDir, err := os.MkdirTemp("", "bz2_extract_")
	if err != nil {
		return []byte{}, fmt.Errorf("创建临时目录失败: %v", err)
	}
	defer os.RemoveAll(tmpDir) // 确保程序退出时清理临时目录
	logger.Logger.Printf("临时目录: %s", tmpDir)

	original := filepath.Base(filename[:len(filename)-len(".bz2")])
	safePath := filepath.Join(tmpDir, sanitizePath(original))

	if err = WriteBz2File(bz2Reader, safePath, os.ModePerm); err != nil {
		return []byte{}, err
	}

	content, cnt, err := WalkDir(tmpDir)
	if err != nil {
		return content, err
	}

	logger.Logger.Printf("bz2文件解析完成，共提取 %d 个文件(一级目录)", cnt)
	return content, nil
}

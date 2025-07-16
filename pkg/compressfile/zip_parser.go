package compressfile

import (
	"archive/zip"
	"fextra/internal"
	"fmt"
	"os"
	"path/filepath"

	"fextra/pkg/logger"
)

type ZipFileParser struct{}

// 提取zip压缩文件中所有文件的内容
func (p *ZipFileParser) Parse(filePath string) ([]byte, error) {
	r, err := zip.OpenReader(filePath)
	if err != nil {
		return []byte{}, fmt.Errorf("无法打开文件: %v", err)
	}
	defer r.Close()
	logger.Logger.Printf("提取文件: %s", filePath)

	// 创建临时目录
	tmpDir, err := os.MkdirTemp("", "zip_extract_")
	if err != nil {
		return []byte{}, fmt.Errorf("创建临时目录失败: %v", err)
	}
	defer os.RemoveAll(tmpDir) // 确保程序退出时清理临时目录
	logger.Logger.Printf("临时目录: %s", tmpDir)

	for _, f := range r.File {
		// 防止路径遍历攻击
		safePath := filepath.Join(tmpDir, sanitizePath(f.Name))

		// 创建目录结构
		if err := os.MkdirAll(filepath.Dir(safePath), 0755); err != nil {
			return []byte{}, fmt.Errorf("创建目录失败 %s: %v", safePath, err)
		}

		logger.DebugLogger.Printf("处理ZIP条目: %s -> %s", f.Name, safePath)
		// 处理目录文件
		if f.FileInfo().IsDir() {
			if err := os.MkdirAll(safePath, 0755); err != nil {
				return []byte{}, fmt.Errorf("创建目录失败 %s: %v", safePath, err)
			}
			continue
		}

		// 打开ZIP内的文件
		rc, err := f.Open()
		if err != nil {
			return []byte{}, fmt.Errorf("打开ZIP内文件 %s 失败: %v", f.Name, err)
		}

		if err := WriteDstFile(rc, safePath, 0755); err != nil {
			rc.Close()
			return []byte{}, fmt.Errorf("写入文件 %s 失败: %v", safePath, err)
		}

		rc.Close()
	}

	content, files, err := walkDir(tmpDir)
	if err != nil {
		return content, err
	}
	logger.Logger.Printf("ZIP文件解析完成，共提取 %d 个文件(一级目录)", files)
	return content, nil
}

func init() {
	// jar(25) zip(21) war(26)
	internal.RegisterParser(internal.FileTypeZIP, &ZipFileParser{})
	internal.RegisterParser(internal.FileTypeJAR, &ZipFileParser{})
	internal.RegisterParser(internal.FileTypeWAR, &ZipFileParser{})
}

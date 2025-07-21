package compressfile

import (
	"archive/tar"
	"bytes"
	"fextra/internal"
	"fextra/pkg/logger"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
)

// Tar header magic number ("ustar\x00\x30\x30") as defined by POSIX standard
const tarMagic = "ustar\x00\x30\x30"

type TarFileParser struct{}

func writeTarFile(tr *tar.Reader, path string, header *tar.Header) error {
	// 创建父目录（如果不存在）
	parentDir := filepath.Dir(path)
	if err := os.MkdirAll(parentDir, 0755); err != nil {
		return err
	}

	// 创建文件
	file, err := os.OpenFile(path, os.O_CREATE|os.O_WRONLY|os.O_TRUNC, os.FileMode(header.Mode))
	if err != nil {
		return err
	}
	defer file.Close()

	// 流式复制内容（避免内存溢出）
	if _, err := io.Copy(file, tr); err != nil {
		return err
	}
	return nil
}

func (p *TarFileParser) Parse(filePath string) ([]byte, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return []byte{}, fmt.Errorf("无法打开文件: %v", err)
	}
	defer file.Close()

	return parseTarFromReader(file)
}

func init() {
	internal.RegisterParser(internal.FileTypeTAR, &TarFileParser{})
}

// parseTarFromReader 从io.Reader解析tar内容并返回格式化字符串
func parseTarFromReader(reader io.Reader) ([]byte, error) {
	tarReader := tar.NewReader(reader)
	var tarContent bytes.Buffer

	// 创建临时目录
	tmpDir, err := os.MkdirTemp("", "tar_extract_")
	if err != nil {
		return tarContent.Bytes(), fmt.Errorf("创建临时目录失败: %v", err)
	}
	defer os.RemoveAll(tmpDir) // 确保程序退出时清理临时目录
	logger.Logger.Printf("临时目录: %s", tmpDir)

	for {
		header, err := tarReader.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			return tarContent.Bytes(), fmt.Errorf("tar解析错误: %v", err)
		}

		targetPath := filepath.Join(tmpDir, sanitizePath(header.Name))

		switch header.Typeflag {
		case tar.TypeDir: // 处理目录
			if err := os.MkdirAll(targetPath, os.FileMode(header.Mode)); err != nil {
				return tarContent.Bytes(), fmt.Errorf("创建目录 %s 失败: %w", targetPath, err)
			}
		case tar.TypeReg: // 处理普通文件
			if err := writeTarFile(tarReader, targetPath, header); err != nil {
				return tarContent.Bytes(), fmt.Errorf("写入文件 %s 失败: %w", targetPath, err)
			}
		}
		logger.Logger.Printf("提取文件: %s", strings.TrimPrefix(targetPath, tmpDir))
	}

	content, files, err := WalkDir(tmpDir)
	if err != nil {
		return content, err
	}
	logger.Logger.Printf("Tar文件解析完成，共提取 %d 个文件(一级目录)", files)
	return content, nil
}

package compressfile

import (
	"bytes"
	"fmt"
	"io"
	"io/fs"
	"os"
	"path/filepath"
	"strings"
	"time"

	"fextra/internal"
	"fextra/pkg/logger"
)

/*
	因为压缩文件内部文件类型不确定，所以在压缩文件解析后，需要再根据文件扩展名选择合适的解压方法
*/

// CompressFileParser 压缩文件解析器
type CompressFileParser struct{}

// GetFullTmpDir 获取完整的临时目录路径
func GetFullTmpDir(tmpdir string) string {
	logger.DebugLogger.Printf("生成临时目录路径，基础路径: %s", tmpdir)
	return filepath.Join(tmpdir, time.Now().Format("20060102150405.000000"))
}

// CreateTmpDir 创建临时目录
func CreateTmpDir(tmpdir string) (string, error) {
	logger.Logger.Printf("开始创建临时目录，基础路径: %s", tmpdir)
	tmpFull := GetFullTmpDir(tmpdir)
	if err := os.MkdirAll(tmpFull, 0755); err != nil {
		logger.Logger.Printf("临时目录创建失败: %v", err)
		return "", err
	}
	logger.Logger.Printf("临时目录创建成功: %s", tmpFull)
	return tmpFull, nil
}

// sanitizePath 防止路径遍历攻击的安全检查
func sanitizePath(path string) string {
	sanitized := strings.TrimPrefix(filepath.Join("/", path), "/")
	if path != sanitized {
		logger.DebugLogger.Printf("路径安全处理: %s -> %s", path, sanitized)
	}
	return sanitized
}

func WriteDstFile(rc io.ReadCloser, safePath string, mode fs.FileMode) error {
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

func walkDir(tmpDir string) ([]byte, int, error) {
	var buffer bytes.Buffer
	var fileCnt int

	err := filepath.Walk(tmpDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if info.IsDir() {
			// filepath.Walk内部实现子目录的递归调用
			return nil
		}

		// 读取文件内容，这里再去校验文件类型，按照对应类型去解析
		fileType := internal.GetDynamicFileType(path)
		parser, err := internal.GetParser(fileType)
		if err != nil {
			return fmt.Errorf("获取解析器失败: %v", err)
		}

		logger.Logger.Printf("walkDir 解析文件: %s", path)
		content, err := parser.Parse(path)
		if err != nil {
			return fmt.Errorf("读取文件 %s 失败: %v", path, err)
		}

		// 在文件解析成功后，添加文件名称等信息
		buffer.WriteString(fmt.Sprintf("=== 文件名: %s ===\n\n", strings.TrimPrefix(path, tmpDir)))
		fileCnt++

		buffer.Write(content)
		buffer.WriteString("\n\n")
		return nil
	})

	return buffer.Bytes(), fileCnt, err
}

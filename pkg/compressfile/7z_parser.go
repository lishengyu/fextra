package compressfile

import (
	"fmt"
	"os"

	"fextra/internal"
	"fextra/pkg/logger"

	"github.com/gen2brain/go-unarr"
)

type SevenZFileParser struct{}

func (p *SevenZFileParser) Parse(filePath string) ([]byte, error) {
	// 打开7z文件
	archive, err := unarr.NewArchive(filePath)
	if err != nil {
		return []byte{}, fmt.Errorf("无法打开7z文件: %v", err)
	}
	defer archive.Close()

	logger.Logger.Printf("提取7z文件: %s", filePath)

	// 创建临时目录
	tmpDir, err := os.MkdirTemp("", "7z_extract_")
	if err != nil {
		return []byte{}, fmt.Errorf("创建临时目录失败: %v", err)
	}
	defer os.RemoveAll(tmpDir) // 确保程序退出时清理临时目录
	logger.Logger.Printf("临时目录: %s", tmpDir)

	files, err := archive.Extract(tmpDir)
	if err != nil {
		return []byte{}, fmt.Errorf("提取7z文件失败: %v", err)
	}
	logger.Logger.Printf("7z文件提取完成，共提取 %d 个文件", len(files))

	// 遍历临时目录并提取所有文件内容
	content, cnt, err := walkDir(tmpDir)
	if err != nil {
		return content, err
	}

	logger.Logger.Printf("7z文件解析完成，共提取 %d 个文件(一级目录)", cnt)
	return content, nil
}

func init() {
	internal.RegisterParser(internal.FileType7Z, &SevenZFileParser{})
	// go-unarr不支持rar v5格式
	//internal.RegisterParser(internal.FileTypeRAR, &SevenZFileParser{})
}

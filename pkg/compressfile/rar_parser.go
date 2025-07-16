package compressfile

import (
	"bytes"
	"fmt"
	"io"
	"os"

	"github.com/nwaples/rardecode"
)

type RarFileParser struct{}

func (p *RarFileParser) Parse(filePath string) ([]byte, error) {
	var content bytes.Buffer

	file, err := os.Open(filePath)
	if err != nil {
		return content.Bytes(), fmt.Errorf("无法打开文件: %v", err)
	}
	defer file.Close()

	reader, err := rardecode.NewReader(file, "") // 空密码
	if err != nil {
		return content.Bytes(), err
	}

	for {
		hdr, err := reader.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			return content.Bytes(), err
		}

		if hdr.IsDir {
			continue
		}

		fmt.Sprintf("=== 文件名: %s ===\n", hdr.Name)

		// 读取所有剩余内容
		remainingData, err := io.ReadAll(reader)
		if err != nil && err != io.EOF {
			return content.Bytes(), err
		}

		if len(remainingData) > 0 {
			content.Write(remainingData)
			content.WriteString("\n\n")
		}
	}

	return content.Bytes(), nil
}

func init() {
	//internal.RegisterParser(internal.FileTypeRAR, &RarFileParser{})
}

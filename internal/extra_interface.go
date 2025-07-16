package internal

import (
	"fmt"
	"os"
)

// FileParser 定义文件解析器接口
type FileParser interface {
	Parse(filePath string) ([]byte, error)
}

var parsers = make(map[int]FileParser)

type UnknownFileParser struct{}

func (p *UnknownFileParser) Parse(filePath string) ([]byte, error) {
	data, err := os.ReadFile(filePath)
	if err != nil {
		return []byte{}, err
	}
	return data, nil
}

// RegisterParser 注册文件类型解析器
func RegisterParser(fileType int, parser FileParser) {
	if _, exists := parsers[fileType]; exists {
		fmt.Printf("警告: 文件类型 %d 已被注册，将忽略重复注册\n", fileType)
		return
	}
	parsers[fileType] = parser
}

// GetParser 获取指定文件类型的解析器
func GetParser(fileType int) (FileParser, error) {
	parser, exists := parsers[fileType]
	if !exists {
		return parsers[114], nil
		//return nil, errors.New("no parser registered for file type")
	}
	return parser, nil
}

func init() {
	RegisterParser(114, &UnknownFileParser{})
}

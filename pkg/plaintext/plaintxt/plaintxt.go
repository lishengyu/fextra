package plaintxt

import "os"

type TextPlainParser struct{}

func (p *TextPlainParser) Parse(filePath string) ([]byte, error) {
	return os.ReadFile(filePath)
}

package main

import (
	"flag"
	"fmt"
	"io"
	"log"
	"os"

	"fextra/internal"
	_ "fextra/pkg/compressfile"
	"fextra/pkg/logger"
	_ "fextra/pkg/office"
)

var (
	InputFile string
	FileType  int
)

func main() {
	flag.StringVar(&InputFile, "i", "", "input file")
	flag.IntVar(&FileType, "t", 0, "file type")
	flag.Parse()
	if InputFile == "" {
		flag.Usage()
		return
	}

	// 启用常规日志输出到控制台
	logger.SetLogger(log.New(os.Stdout, "[Fextra Logger] ", log.LstdFlags))
	// 启用调试日志
	logger.DebugLogger = log.New(io.Discard, "", 0)
	//logger.SetDebugLogger(log.New(os.Stdout, "[Fextra Logger Debug] ", log.LstdFlags))

	if FileType == 0 {
		// 动态获取文件类型
		FileType = internal.GetDynamicFileType(InputFile)
	}

	parser, err := internal.GetParser(FileType)
	if err != nil {
		fmt.Println(err)
		return
	}

	text, err := parser.Parse(InputFile)
	if err != nil {
		fmt.Println(err)
		return
	}

	fmt.Printf("file[%s], size[%d], content:%s\n", InputFile, len(text), string(text))
}

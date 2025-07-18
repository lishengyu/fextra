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
	InputFile     string
	FileType      int
	Verbose       bool
	DetailVerbose bool
)

func main() {
	flag.StringVar(&InputFile, "i", "", "input file")
	flag.IntVar(&FileType, "t", 0, "file type")
	flag.BoolVar(&Verbose, "v", false, "verbose")
	flag.BoolVar(&DetailVerbose, "vv", false, "detail verbose")

	flag.Parse()
	if InputFile == "" {
		flag.Usage()
		return
	}

	if DetailVerbose {
		// 启用常规日志输出到控制台
		logger.SetLogger(log.New(os.Stdout, "[Fextra Logger] ", log.LstdFlags))
		// 启用调试日志
		logger.SetDebugLogger(log.New(os.Stdout, "[Fextra Logger Debug] ", log.LstdFlags))
	} else if Verbose {
		// 启用常规日志输出到控制台
		logger.SetLogger(log.New(os.Stdout, "[Fextra Logger] ", log.LstdFlags))
		// 启用调试日志
		logger.DebugLogger = log.New(io.Discard, "", 0)
	}

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

	logger.Logger.Printf("content:\n%s\n", string(text))
	fmt.Printf("file[%s], size[%d]\n", InputFile, len(text))
}

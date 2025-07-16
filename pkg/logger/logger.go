package logger

import (
	"io"
	"log"
)

// StdLogger 日志接口定义
type StdLogger interface {
	Print(v ...interface{})
	Printf(format string, v ...interface{})
	Println(v ...interface{})
}

var (
	// Logger 用于记录常规日志，默认丢弃所有日志
	Logger StdLogger = log.New(io.Discard, "[Fextra] ", log.LstdFlags)

	// DebugLogger 用于记录调试日志，默认使用Logger
	DebugLogger StdLogger = &debugLogger{}
)

// debugLogger 调试日志转发器
type debugLogger struct{}

func (d *debugLogger) Print(v ...interface{}) {
	Logger.Print(v...)
}
func (d *debugLogger) Printf(format string, v ...interface{}) {
	Logger.Printf(format, v...)
}
func (d *debugLogger) Println(v ...interface{}) {
	Logger.Println(v...)
}

// SetLogger 设置全局日志实例
func SetLogger(l StdLogger) {
	Logger = l
}

// SetDebugLogger 设置调试日志实例
func SetDebugLogger(l StdLogger) {
	DebugLogger = l
}

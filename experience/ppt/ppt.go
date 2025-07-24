package ppt

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fextra/pkg/logger"
	"fmt"
	"os"
	"strings"
	"unicode/utf16"

	"github.com/richardlehane/mscfb"
)

// 简化实现：提取所有可能的文本片段
// 规范3.3章节中提到了powerpoint document stream是record类型集合，
// 具体类型由RecordHeader进行标识

const (
	RecordHeaderLen = 8
	// 文本相关记录类型
	RT_TextBytesAtom = 0x0FA8
	RT_TextCharsAtom = 0x0FA0
	RT_CStringAtom   = 0x0FBA

	RT_TextHeaderAtom   = 0x003F
	RT_TextSpecInfoAtom = 0x0040
	RT_TextRulerAtom    = 0x0050
	RT_TextStyleAtom    = 0x0053
)

var (
	// 文本记录类型集合
	extTextRecordTypes = map[uint16]bool{
		//0x0FA0: true, 0x0FA1: true, 0x0FA2: true, 0x0FA3: true, 0x0FA4: true,
		//0x0FA5: true, 0x0FA6: true, 0x0FA7: true, 0x0FA8: true, 0x0FA9: true,
		//0x0FAA: true, 0x0FAB: true, 0x0FAC: true, 0x0FAD: true, 0x0FAE: true,
		//0x0FAF: true, 0x0FF6: true,
		RT_TextBytesAtom: true, RT_TextCharsAtom: true, RT_CStringAtom: true,
	}
)

// PPTNode 表示PPT记录树中的节点
// 容器记录(Container Record)会包含子节点，原子记录(Atom Record)为叶子节点
type PPTNode struct {
	Header   RecordHeader // 记录头信息
	Data     []byte       // 记录数据
	Parent   *PPTNode     // 父节点
	Children []*PPTNode   // 子节点列表
}

type PptParse struct {
	File              *mscfb.Reader
	PptDocumentStream []byte
	StreamLen         int      // PptDocumentStream的大小
	StreamOffset      int      // PptDocumentStream的当前偏移
	RecordNum         int      // PptDocumentStream的记录数量
	RootNode          *PPTNode // 记录树的根节点
	CurrentNode       *PPTNode // 当前解析节点
}

type OfficePptParser struct{}

// 解码UTF-16字节流为字符串
func decodeUTF16(data []byte, byteOrder binary.ByteOrder) string {
	var bomSize int
	if len(data) >= 2 {
		if data[0] == 0xFF && data[1] == 0xFE {
			byteOrder = binary.LittleEndian
			bomSize = 2
		} else if data[0] == 0xFE && data[1] == 0xFF {
			byteOrder = binary.BigEndian
			bomSize = 2
		}
	}
	if byteOrder == nil {
		byteOrder = binary.LittleEndian
	}

	u16s := make([]uint16, (len(data)-bomSize)/2)
	for i := 0; i < len(u16s); i++ {
		u16s[i] = byteOrder.Uint16(data[bomSize+2*i:])
	}

	var runes []rune
	for i := 0; i < len(u16s); {
		if utf16.IsSurrogate(rune(u16s[i])) && i+1 < len(u16s) {
			r := utf16.DecodeRune(rune(u16s[i]), rune(u16s[i+1]))
			runes = append(runes, r)
			i += 2
		} else {
			runes = append(runes, rune(u16s[i]))
			i++
		}
	}
	return string(runes)
}

func NewPptParse(file *os.File) (*PptParse, error) {
	doc, err := mscfb.New(file)
	if err != nil {
		return nil, fmt.Errorf("文件打开失败: %w", err)
	}
	return &PptParse{
		File:              doc,
		PptDocumentStream: make([]byte, 0),
		StreamLen:         0,
		StreamOffset:      0,
		RootNode:          &PPTNode{}, // 初始化根节点
		CurrentNode:       nil,
	}, nil
}

func (d *PptParse) GetPptDocumentStream() error {
	if d.File == nil {
		return errors.New("mscfb file is nil")
	}

	var buf []byte
	for _, file := range d.File.File {
		logger.Logger.Printf("file name: %s", file.Name)
		if file.Name == "PowerPoint Document" {
			buf = make([]byte, file.Size)
			n, err := file.Read(buf)
			if err != nil {
				return fmt.Errorf("failed to open PowerPoint Document stream: %v", err)
			}
			logger.Logger.Printf("read %d bytes： %v", n, buf[:32])
			d.PptDocumentStream = buf
			d.StreamLen = n
			return nil
		}
	}

	return fmt.Errorf("PowerPoint Document stream not found")
}

func (d *PptParse) parseTextRecords() ([]byte, error) {
	if len(d.PptDocumentStream) == 0 {
		return nil, errors.New("PPT文档流为空")
	}

	var textBuffer bytes.Buffer
	// 从根节点开始解析记录树
	d.CurrentNode = d.RootNode
	if _, err := d.parseRecordToNode(&textBuffer); err != nil {
		return textBuffer.Bytes(), err
	}

	// 可以添加树的遍历逻辑，用于调试或其他处理
	// d.traverseNode(d.RootNode, 0)

	return textBuffer.Bytes(), nil
}

// 遍历节点树，用于调试或打印结构
func (d *PptParse) traverseNode(node *PPTNode, depth int) {
	indent := strings.Repeat("  ", depth)
	logger.Logger.Printf("%s类型: 0x%04x, 版本: 0x%x, 实例: 0x%x, 子节点数: %d",
		indent, node.Header.RecType, node.Header.RecVer, node.Header.RecInstance, len(node.Children))
	for _, child := range node.Children {
		d.traverseNode(child, depth+1)
	}
}

func (d *PptParse) ExtractText() ([]byte, error) {
	if err := d.GetPptDocumentStream(); err != nil {
		return nil, err
	}

	return d.parseTextRecords()
}

func (p *OfficePptParser) Parse(filePath string) ([]byte, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, fmt.Errorf("文件打开失败: %w", err)
	}
	defer file.Close()

	parser, err := NewPptParse(file)
	if err != nil {
		return nil, fmt.Errorf("初始化PPT解析器失败: %w", err)
	}

	content, err := parser.ExtractText()
	if err != nil {
		return content, fmt.Errorf("提取文本失败: %w", err)
	}

	return content, nil
}

// RecordHeader 表示PPT记录头结构，遵循MS-PPT规范2.3.1节定义
type RecordHeader struct {
	RecVer      uint8  // 4位: 记录版本，0xF表示容器记录
	RecInstance uint16 // 12位: 记录实例数据
	RecType     uint16 // 2字节: 记录类型
	RecLen      uint32 // 4字节: 记录数据长度
}

// parseRecordHeader 解析8字节的记录头结构
// 返回解析后的RecordHeader、新的位置偏移和可能的错误
func parseRecordHeader(stream []byte, pos int) (RecordHeader, int, error) {
	// 检查流边界
	if pos+RecordHeaderLen > len(stream) {
		return RecordHeader{}, pos, fmt.Errorf("记录头超出流边界，需要%d字节，剩余%d字节", RecordHeaderLen, len(stream)-pos)
	}

	// 读取前2字节(16位)，包含RecVer和RecInstance
	verInstance := binary.LittleEndian.Uint16(stream[pos : pos+2])

	// 解析4位RecVer和12位RecInstance
	// 低4位为RecVer，高12位为RecInstance
	recVer := uint8(verInstance & 0x000F)   // 低4位为RecVer
	recInstance := uint16(verInstance >> 4) // 高12位为RecInstance

	return RecordHeader{
		RecVer:      recVer,
		RecInstance: recInstance,
		RecType:     binary.LittleEndian.Uint16(stream[pos+2 : pos+4]),
		RecLen:      binary.LittleEndian.Uint32(stream[pos+4 : pos+8]),
	}, pos + RecordHeaderLen, nil
}

func (d *PptParse) parseContainer(textBuffer *bytes.Buffer, offset int, recordEnd int) error {
	for d.StreamOffset < recordEnd { // 递归解析子记录
		if err := d.parseRecord(textBuffer); err != nil {
			return fmt.Errorf("解析子记录失败: %w", err)
		}
	}
	return nil
}

func (d *PptParse) parseRecord(textBuffer *bytes.Buffer) error {
	// 创建根节点并开始解析
	d.CurrentNode = d.RootNode
	_, err := d.parseRecordToNode(textBuffer)
	return err
}

func (d *PptParse) parseRecordToNode(textBuffer *bytes.Buffer) (*PPTNode, error) {
	stream := d.PptDocumentStream

	for d.StreamOffset+RecordHeaderLen < d.StreamLen {
		// 解析记录头
		header, newPos, err := parseRecordHeader(stream, d.StreamOffset)
		if err != nil {
			return nil, fmt.Errorf("解析记录头失败: %w", err)
		}
		d.StreamOffset = newPos // 解析完header后的偏移

		recordEnd := d.StreamOffset + int(header.RecLen) // 当前record的结束位置

		// 创建新节点
		node := &PPTNode{
			Header: header,
			Parent: d.CurrentNode,
		}

		// 边界检查
		if recordEnd > len(stream) {
			return nil, fmt.Errorf("记录超出流边界，类型: 0x%04x, 版本: 0x%x, 预期长度: %d, 剩余字节: %d",
				header.RecType, header.RecVer, header.RecLen, d.StreamLen-d.StreamOffset)
		}

		d.RecordNum++
		logger.Logger.Printf("解析第%d条记录, stream偏移：0x%x, 类型: 0x%04x, 版本: 0x%x, 长度: 0x%x字节",
			d.RecordNum, d.StreamOffset, header.RecType, header.RecVer, header.RecLen)

		// 读取记录数据
		node.Data = stream[d.StreamOffset:recordEnd]

		// 1. 处理容器记录（如RT_Document=0x03E8）
		if header.RecVer == 0xF { // 容器记录由RecVer=0xF标识
			// 递归解析子记录
			if err := d.parseContainer(textBuffer, d.StreamOffset, recordEnd); err != nil {
				return nil, fmt.Errorf("解析容器记录失败: %w", err)
			}
		} else if extTextRecordTypes[header.RecType] {
			// 2. 处理文本记录
			text := decodeUTF16(node.Data, binary.LittleEndian)
			text = strings.TrimSpace(text)

			logger.DebugLogger.Printf("解析文本记录, stream偏移：0x%x, 类型: 0x%04x, 版本: 0x%x, 长度: 0x%x字节, 文本内容: %s",
				d.StreamOffset, header.RecType, header.RecVer, header.RecLen, text)
			if text != "" {
				textBuffer.WriteString(fmt.Sprintf("=== 文本内容 ===\n%s\n\n", text))
			}
		} else {
			logger.DebugLogger.Printf("忽略未知记录类型: 0x%04x, stream偏移：0x%x,版本: 0x%x, 长度: 0x%x字节",
				header.RecType, d.StreamOffset, header.RecVer, header.RecLen)
		}

		d.StreamOffset = recordEnd
	}
	return nil, nil
}

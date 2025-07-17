package xls

import (
	"bytes"
	"encoding/binary"
	"encoding/hex"
	"errors"
	"fextra/internal"
	"fextra/pkg/logger"
	"fmt"
	"os"
	"strings"
	"unicode/utf16"
)

const (
	XlsSignature    = "d0cf11e0a1b11ae1"
	XlsHeaderOffset = 512
)

// 文件头结构 (512字节)
type FileHeader struct {
	Signature            [8]byte     // 文件标识：0xD0CF11E0A1B11AE1
	CLSID                [16]byte    // 保留字段
	MinorVersion         uint16      // 次要版本
	MajorVersion         uint16      // 主要版本（3或4）
	ByteOrder            uint16      // 字节序（0xFFFE为小端序）
	SectorShift          uint16      // 扇区大小（512=0x0009, 4096=0x000C）
	MiniSectorShift      uint16      // 迷你扇区大小（固定64字节 = 0x0006）
	Reserved             [6]byte     // 保留字段
	DirectorySectorCnt   uint32      // 目录扇区数量,MajorVersion=3时为0
	FATSectorCnt         uint32      // FAT表扇区数量
	DirectoryStart       uint32      // 目录起始扇区ID
	TransactionSignature uint32      // 事务签名（MajorVersion=4时使用）
	MiniStreamCutoffSize uint32      // 迷你流截断大小
	MiniFATStart         uint32      // 迷你FAT起始扇区ID
	MiniFATSectorCnt     uint32      // 迷你FAT扇区数量
	DiFATSectorStart     uint32      // DIFAT起始扇区ID
	DIFATSectorCnt       uint32      // DIFAT扇区数量
	DiFAT                [109]uint32 // DIFAT扇区ID数组
}

// 目录项结构 (128字节)
type DirectoryEntry struct {
	Name           [64]byte // UTF-16名称
	NameLen        uint16   // 实际名称长度
	ObjectType     uint8    // 类型：0x0(unknown) 0x01(存储) 0x02(流) 0x05(根存储)
	ColorFlag      uint8    // 颜色标志
	LeftSiblingID  uint32   // 左兄弟项ID
	RightSiblingID uint32   // 右兄弟项ID
	ChildID        uint32   // 子项ID
	CLSID          [16]byte // CLSID
	StateBits      uint32   // 状态位
	CreationTime   int64    // 创建时间
	ModifiedTime   int64    // 修改时间
	StartSectorID  uint32   // 流起始扇区ID
	StreamSize     uint64   // 流大小
}

type XDirectoryEntry struct {
	Name  string
	Type  uint8
	Entry *DirectoryEntry
}

type XlsParse struct {
	File *os.File

	FileHeader *FileHeader
	DirEntry   []*XDirectoryEntry
	FAT        []uint32
	MiniFAT    []uint32

	WorkbookStream []byte
	SectorSize     int

	// XLS特定字段
	BofSectorStartID uint32
	BofSectorSize    uint64
}

type OfficeXlsParser struct{}

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

	return string(utf16.Decode(u16s))
}

func (h *FileHeader) Printf() {
	logger.Logger.Printf("XLS文件版本: %d.%d", h.MajorVersion, h.MinorVersion)
	logger.Logger.Printf("扇区大小: %d, 扇区数量: %d", 1<<h.SectorShift, h.FATSectorCnt)
}

func (e *XDirectoryEntry) CheckWorkbookStream() bool {
	return e.Type == 0x02 && strings.EqualFold(e.Name, "workbook")
}

func NewXlsParse(fn string) (*XlsParse, error) {
	file, err := os.Open(fn)
	if err != nil {
		return nil, fmt.Errorf("文件打开失败: %w", err)
	}
	return &XlsParse{
		File:           file,
		FileHeader:     &FileHeader{},
		DirEntry:       make([]*XDirectoryEntry, 0),
		FAT:            make([]uint32, 0),
		MiniFAT:        make([]uint32, 0),
		WorkbookStream: make([]byte, 0),
	}, nil
}

func (d *XlsParse) Close() {
	if d.File != nil {
		d.File.Close()
	}
}

func (d *XlsParse) ParseHeader() error {
	header := &FileHeader{}
	if err := binary.Read(d.File, binary.LittleEndian, header); err != nil {
		return err
	}

	if hex.EncodeToString(header.Signature[:]) != XlsSignature {
		return errors.New("无效的XLS OLE签名")
	}

	header.Printf()
	d.SectorSize = 1 << header.SectorShift
	d.FileHeader = header
	return nil
}

func (d *XlsParse) LoadFAT() error {
	entriesPerSector := d.SectorSize / 4
	fat := make([]uint32, 0, d.FileHeader.FATSectorCnt*uint32(entriesPerSector))

	for _, fatSectorID := range d.FileHeader.DiFAT {
		if fatSectorID == 0xFFFFFFFF {
			continue
		}
		sectorPos := int64(XlsHeaderOffset) + int64(fatSectorID)*int64(d.SectorSize)
		if _, err := d.File.Seek(sectorPos, 0); err != nil {
			return err
		}

		entries := make([]uint32, entriesPerSector)
		if err := binary.Read(d.File, binary.LittleEndian, &entries); err != nil {
			return err
		}
		fat = append(fat, entries...)
	}
	d.FAT = fat
	return nil
}

func (d *XlsParse) GetDirEntries() error {
	dirSectorPos := int64(XlsHeaderOffset) + int64(d.FileHeader.DirectoryStart)*int64(d.SectorSize)
	if _, err := d.File.Seek(dirSectorPos, 0); err != nil {
		return err
	}

	direntryCount := d.SectorSize / 128
	if d.FileHeader.MajorVersion > 3 {
		direntryCount = int(d.FileHeader.DirectorySectorCnt+1) * direntryCount
	}

	for i := 0; i < direntryCount; i++ {
		entry := &DirectoryEntry{}
		if err := binary.Read(d.File, binary.LittleEndian, entry); err != nil {
			break
		}

		if entry.NameLen == 0 {
			continue
		}

		name := decodeUTF16(entry.Name[:entry.NameLen], binary.LittleEndian)
		xd := &XDirectoryEntry{
			Name:  name,
			Type:  entry.ObjectType,
			Entry: entry,
		}
		d.DirEntry = append(d.DirEntry, xd)

		if xd.CheckWorkbookStream() {
			d.BofSectorStartID = entry.StartSectorID
			d.BofSectorSize = entry.StreamSize
			logger.Logger.Printf("找到Workbook流: %s, 起始扇区: %d, 大小: %d", name, entry.StartSectorID, entry.StreamSize)
		}
	}

	if len(d.DirEntry) == 0 {
		return errors.New("未找到目录项")
	}
	return nil
}

func (d *XlsParse) GetWorkbookStream() error {
	if d.BofSectorStartID == 0 || d.BofSectorSize == 0 {
		return errors.New("未找到有效的Workbook流")
	}

	var buffer bytes.Buffer
	currentSector := d.BofSectorStartID
	pos := uint64(0)

	for currentSector != 0xFFFFFFFE && pos < d.BofSectorSize {
		sectorPos := int64(XlsHeaderOffset) + int64(currentSector)*int64(d.SectorSize)
		if _, err := d.File.Seek(sectorPos, 0); err != nil {
			return err
		}

		readSize := uint64(d.SectorSize)
		if pos+readSize > d.BofSectorSize {
			readSize = d.BofSectorSize - pos
		}

		data := make([]byte, readSize)
		if _, err := d.File.Read(data); err != nil {
			return err
		}

		buffer.Write(data)
		pos += readSize
		currentSector = d.FAT[currentSector]
	}

	d.WorkbookStream = buffer.Bytes()
	logger.Logger.Printf("Workbook流提取完成，大小: %d字节", len(d.WorkbookStream))
	return nil
}

func (d *XlsParse) parseCellRecords() ([]byte, error) {
	// XLS二进制格式单元格记录解析逻辑
	var textBuffer bytes.Buffer
	stream := d.WorkbookStream
	pos := 0

	// XLS文本通常存储在类型为0x0204(LabelSst)和0x00FD(RString)的记录中
	for pos+4 < len(stream) {
		recordType := binary.LittleEndian.Uint16(stream[pos:])
		recordLen := binary.LittleEndian.Uint16(stream[pos+2:])
		pos += 4

		// 检查文本相关记录类型
		if recordType == 0x0204 || recordType == 0x00FD {
			if pos+int(recordLen) > len(stream) {
				break
			}

			// 提取并解码文本内容
			textData := stream[pos : pos+int(recordLen)]
			text := decodeUTF16(textData, binary.LittleEndian)
			text = strings.TrimSpace(text)

			if text != "" {
				textBuffer.WriteString(fmt.Sprintf("%s\n", text))
			}
		}

		pos += int(recordLen)
	}

	return textBuffer.Bytes(), nil
}

func (d *XlsParse) ExtractText() ([]byte, error) {
	if err := d.GetWorkbookStream(); err != nil {
		return nil, err
	}

	return d.parseCellRecords()
}

func (p *OfficeXlsParser) Parse(filePath string) ([]byte, error) {
	parser, err := NewXlsParse(filePath)
	if err != nil {
		return nil, fmt.Errorf("初始化XLS解析器失败: %w", err)
	}
	defer parser.Close()

	if err := parser.ParseHeader(); err != nil {
		return nil, fmt.Errorf("解析文件头失败: %w", err)
	}

	if err := parser.LoadFAT(); err != nil {
		return nil, fmt.Errorf("加载FAT表失败: %w", err)
	}

	if err := parser.GetDirEntries(); err != nil {
		return nil, fmt.Errorf("获取目录项失败: %w", err)
	}

	content, err := parser.ExtractText()
	if err != nil {
		return content, fmt.Errorf("提取文本失败: %w", err)
	}

	return content, nil
}

func init() {
	internal.RegisterParser(internal.FileTypeXLS, &OfficeXlsParser{})
}

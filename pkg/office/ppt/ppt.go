package ppt

import (
	"bytes"
	"encoding/binary"
	"encoding/hex"
	"errors"
	"fextra/pkg/logger"
	"fmt"
	"os"
	"strings"
	"unicode/utf16"
)

const (
	PptSignature    = "d0cf11e0a1b11ae1"
	PptHeaderOffset = 512
)

// 文件头结构 (512字节) - 与DOC格式相同
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

type PDirectoryEntry struct {
	Name  string
	Type  uint8
	Entry *DirectoryEntry
}

type PptParse struct {
	File *os.File

	FileHeader *FileHeader
	DirEntry   []*PDirectoryEntry
	DIFAT      []uint32
	FAT        []uint32
	MiniFAT    []uint32

	PptDocumentStream []byte
	SectorSize        int

	// PPT特定字段
	SlideSectorStartID uint32
	SlideSectorSize    uint64
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

func (h *FileHeader) Printf() {
	logger.Logger.Printf("文件版本:     %d.%d\n", h.MajorVersion, h.MinorVersion)
	logger.Logger.Printf("扇区大小：    %d,  扇区数量:     %d\n", 1<<h.SectorShift, h.FATSectorCnt)
	logger.Logger.Printf("迷你扇区大小：%d,  迷你扇区数量：%d, 迷你扇区起始ID：%d\n", 1<<h.MiniSectorShift, h.MiniFATSectorCnt, h.MiniFATStart)
	logger.Logger.Printf("目录扇区数量：%d   目录扇区起始ID：%d\n", h.DirectorySectorCnt, h.DirectoryStart)
	logger.Logger.Printf("Di目录项数量：%d,  Di目录项起始ID：%d\n", h.DIFATSectorCnt, h.DiFATSectorStart)
}

func (e *PDirectoryEntry) CheckPptDocumentStream() bool {
	return e.Type == 0x02 && strings.Contains(strings.ToLower(e.Name), "powerpoint document")
}

func NewPptParse(fn string) (*PptParse, error) {
	file, err := os.Open(fn)
	if err != nil {
		return nil, fmt.Errorf("文件打开失败: %w", err)
	}
	return &PptParse{
		File:              file,
		FileHeader:        &FileHeader{},
		DirEntry:          make([]*PDirectoryEntry, 0),
		DIFAT:             make([]uint32, 0),
		FAT:               make([]uint32, 0),
		MiniFAT:           make([]uint32, 0),
		PptDocumentStream: make([]byte, 0),
	}, nil
}

func (d *PptParse) Close() {
	if d.File != nil {
		d.File.Close()
	}
}

func (d *PptParse) ParseHeader() error {
	header := &FileHeader{}
	if err := binary.Read(d.File, binary.LittleEndian, header); err != nil {
		return err
	}

	if hex.EncodeToString(header.Signature[:]) != PptSignature {
		return errors.New("无效的PPT OLE签名")
	}

	header.Printf()
	d.SectorSize = 1 << header.SectorShift
	d.FileHeader = header
	return nil
}

func (d *PptParse) LoadDIFAT() error {
	header := d.FileHeader
	file := d.File

	// 1. 处理头部109个DIFAT条目
	difat := make([]uint32, 0, 109+int(header.DIFATSectorCnt)*d.SectorSize/4)
	for _, sector := range header.DiFAT {
		if sector != 0xFFFFFFFF { // 0xFFFFFFFF表示空条目
			difat = append(difat, sector)
		}
	}

	// 2. 处理额外的DIFAT扇区
	currentSector := header.DiFATSectorStart
	for i := uint32(0); i < header.DIFATSectorCnt; i++ {
		sectorPos := PptHeaderOffset + int64(currentSector)*int64(d.SectorSize)
		_, err := file.Seek(sectorPos, 0)
		if err != nil {
			return err
		}

		// 每个DIFAT扇区包含 (扇区大小/4 - 1) 个条目
		entries := make([]uint32, d.SectorSize/4-1)
		if err := binary.Read(file, binary.LittleEndian, &entries); err != nil {
			return err
		}

		// 读取下一个DIFAT扇区指针（位于扇区末尾）
		var nextSector uint32
		if err := binary.Read(file, binary.LittleEndian, &nextSector); err != nil {
			return err
		}

		difat = append(difat, entries...)
		currentSector = nextSector
	}

	d.DIFAT = difat // 存储DIFAT扇区ID列表
	// 指示哪些扇区是FAT表，用于FAT表内容的读取
	logger.Logger.Printf("DiFAT扇区表： %v\n", difat)
	return nil
}

func (d *PptParse) LoadFAT() error {
	entriesPerSector := d.SectorSize / 4
	fat := make([]uint32, 0, d.FileHeader.FATSectorCnt*uint32(entriesPerSector))

	for _, fatSectorID := range d.DIFAT {
		if fatSectorID == 0xFFFFFFFF {
			continue
		}
		sectorPos := int64(PptHeaderOffset) + int64(fatSectorID)*int64(d.SectorSize)
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
	logger.Logger.Printf("FAT表数量: %d,  扇区数量: %d\n", len(fat), d.FileHeader.FATSectorCnt)
	return nil
}

func (d *PptParse) LoadMiniFAT() error {
	header := d.FileHeader
	file := d.File

	if header.MiniFATSectorCnt == 0 {
		// 没有MiniFAT
		return nil
	}

	sectorNum := header.MiniFATSectorCnt
	currentSector := header.MiniFATStart
	miniFAT := make([]uint32, header.MiniFATSectorCnt*(uint32(d.SectorSize)/4)) //每个条目4字节
	logger.Logger.Printf("Mini扇区 ====> 数量：%d  大小: %d, 起始分区id: %d\n", sectorNum, d.SectorSize, currentSector)

	sectorPos := int64(512 + int(currentSector)*d.SectorSize)
	logger.Logger.Printf("Mini扇区起始偏移: 0x%x\n", sectorPos)

	_, err := file.Seek(sectorPos, 0)
	if err != nil {
		return err
	}

	// 读取Mini FAT表（每个条目4字节）
	for i := range miniFAT {
		if err := binary.Read(file, binary.LittleEndian, &miniFAT[i]); err != nil {
			return err
		}
	}
	d.MiniFAT = miniFAT
	logger.DebugLogger.Printf("迷你扇区细节[%d]： %v\n", len(miniFAT), miniFAT)
	return nil
}

func (d *PptParse) GetDirentryCount() int {
	var direntryCount int
	if d.FileHeader.MajorVersion == 3 {
		currentSector := d.FileHeader.DirectoryStart
		for currentSector != 0xFFFFFFFE {
			currentSector = d.FAT[currentSector]
			direntryCount += d.SectorSize / 128
		}
	} else {
		direntryCount = int(d.FileHeader.DirectorySectorCnt+1) * direntryCount
	}

	logger.Logger.Printf("目录项数量: %d\n", direntryCount)
	return direntryCount
}

func (d *PptParse) GetDirEntries() error {
	dirSectorPos := int64(PptHeaderOffset) + int64(d.FileHeader.DirectoryStart)*int64(d.SectorSize)
	if _, err := d.File.Seek(dirSectorPos, 0); err != nil {
		return err
	}

	direntryCount := d.GetDirentryCount()

	for i := 0; i < direntryCount; i++ {
		entry := &DirectoryEntry{}
		if err := binary.Read(d.File, binary.LittleEndian, entry); err != nil {
			break
		}

		if entry.NameLen == 0 {
			continue
		}

		name := decodeUTF16(entry.Name[:entry.NameLen], binary.LittleEndian)
		pd := &PDirectoryEntry{
			Name:  name,
			Type:  entry.ObjectType,
			Entry: entry,
		}
		d.DirEntry = append(d.DirEntry, pd)

		if pd.CheckPptDocumentStream() {
			d.SlideSectorStartID = entry.StartSectorID
			d.SlideSectorSize = entry.StreamSize
			logger.Logger.Printf("找到PPT文档流: %s, 起始扇区: %d, 大小: %d", name, entry.StartSectorID, entry.StreamSize)
		}

		logger.Logger.Printf("目录项名称: %s, 长度： %d, 类型: %d, 起始扇区: %d, 大小: %d\n",
			name, entry.NameLen, entry.ObjectType, entry.StartSectorID, entry.StreamSize)
	}

	if len(d.DirEntry) == 0 {
		return errors.New("未找到目录项")
	}
	return nil
}

func (d *PptParse) GetPptDocumentStream() error {
	if d.SlideSectorStartID == 0 || d.SlideSectorSize == 0 {
		return errors.New("未找到有效的PPT文档流")
	}

	var buffer bytes.Buffer
	currentSector := d.SlideSectorStartID
	pos := uint64(0)

	for currentSector != 0xFFFFFFFE && pos < d.SlideSectorSize {
		sectorPos := int64(PptHeaderOffset) + int64(currentSector)*int64(d.SectorSize)
		if _, err := d.File.Seek(sectorPos, 0); err != nil {
			return err
		}

		readSize := uint64(d.SectorSize)
		if pos+readSize > d.SlideSectorSize {
			readSize = d.SlideSectorSize - pos
		}

		data := make([]byte, readSize)
		if _, err := d.File.Read(data); err != nil {
			return err
		}

		buffer.Write(data)
		pos += readSize
		currentSector = d.FAT[currentSector]
	}

	d.PptDocumentStream = buffer.Bytes()
	logger.Logger.Printf("PPT文档流提取完成，大小: %d字节", len(d.PptDocumentStream))
	return nil
}

// 涉及到递归调用，解析嵌套的record记录
func (d *PptParse) parseRecord(stream []byte, pos *int, textBuffer *bytes.Buffer) {
	for *pos+RecordHeaderLen < len(stream) {
		recordVer := binary.LittleEndian.Uint16(stream[*pos:])
		recordType := binary.LittleEndian.Uint16(stream[*pos+2:])
		recordLen := binary.LittleEndian.Uint32(stream[*pos+4:])
		recordEnd := *pos + RecordHeaderLen + int(recordLen)
		tmpCount++
		logger.Logger.Printf("stream偏移：%d, 解析第%d条记录,记录版本: 0x%x, 当前记录类型: 0x%x, 记录长度: %d",
			*pos, tmpCount, recordVer, recordType, recordLen)

		// 1. 处理容器记录（如RT_Document=0x03E8）
		if (recordType >= 0x0F00 && recordType <= 0x0FFF) || recordType == 0x03E8 {
			// 递归解析子记录
			subPos := *pos + RecordHeaderLen
			for subPos < recordEnd {
				d.parseRecord(stream, &subPos, textBuffer)
			}
		} else if (recordType >= 0x0FA0 && recordType <= 0x0FAF) || recordType == 0x0FF6 {
			// PPT文本通常存储在类型为0x0FA0-0x0FAF的记录中
			if *pos+int(recordLen) > len(stream) {
				break
			}
			logger.Logger.Printf("记录类型: %x, 文本记录偏移：%d, 文本记录长度: %d", recordType, *pos, recordLen)
			// 提取并解码文本内容
			textData := stream[*pos : *pos+int(recordLen)]
			text := decodeUTF16(textData, binary.LittleEndian)
			text = strings.TrimSpace(text)

			if text != "" {
				textBuffer.WriteString(fmt.Sprintf("=== 文本内容 ===\n%s\n\n", text))
			}
		}

		*pos = recordEnd
	}
}

/*
PPT二进制格式文本记录解析逻辑

	简化实现：提取所有可能的文本片段
	规范3.3章节中提到了powerpoint document stream是record类型集合，
	具体类型由RecordHeader进行标识
*/

const (
	RecordHeaderLen = 8
)

var (
	tmpCount = 0
)

func (d *PptParse) parseTextRecords() ([]byte, error) {
	var textBuffer bytes.Buffer
	stream := d.PptDocumentStream
	pos := 0

	// 解析records序列，同时records序列中可能存在嵌套的record记录
	d.parseRecord(stream, &pos, &textBuffer)

	return textBuffer.Bytes(), nil
}

func (d *PptParse) ExtractText() ([]byte, error) {
	if err := d.GetPptDocumentStream(); err != nil {
		return nil, err
	}

	return d.parseTextRecords()
}

func (p *OfficePptParser) Parse(filePath string) ([]byte, error) {
	parser, err := NewPptParse(filePath)
	if err != nil {
		return nil, fmt.Errorf("初始化PPT解析器失败: %w", err)
	}
	defer parser.Close()

	if err = parser.ParseHeader(); err != nil {
		return nil, fmt.Errorf("解析文件头失败: %w", err)
	}

	if err = parser.LoadDIFAT(); err != nil {
		return nil, fmt.Errorf("加载DIFAT表失败: %w", err)
	}

	if err = parser.LoadFAT(); err != nil {
		return nil, fmt.Errorf("加载FAT表失败: %w", err)
	}

	if err = parser.LoadMiniFAT(); err != nil {
		return nil, fmt.Errorf("加载MiniFAT表失败: %w", err)
	}

	if err = parser.GetDirEntries(); err != nil {
		return nil, fmt.Errorf("获取目录项失败: %w", err)
	}

	content, err := parser.ExtractText()
	if err != nil {
		return content, fmt.Errorf("提取文本失败: %w", err)
	}

	return content, nil
}

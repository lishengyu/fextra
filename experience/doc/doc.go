package doc

import (
	"bytes"
	"encoding/binary"
	"encoding/hex"
	"errors"
	"fextra/pkg/logger"
	"fextra/pkg/office/doc/fib"
	"fmt"
	"os"
	"strings"
	"unicode/utf16"
	"unicode/utf8"

	"golang.org/x/text/encoding/simplifiedchinese"
)

const (
	DocSignature    = "d0cf11e0a1b11ae1"
	DocHeaderOffset = 512
)

// 文件头结构 (512字节)
type FileHeader struct {
	Signature            [8]byte     // 文件标识：0xD0CF11E0A1B11AE1 [1,8](@ref)
	CLSID                [16]byte    // 保留字段
	MinorVersion         uint16      // 次要版本
	MajorVersion         uint16      // 主要版本（3或4）
	ByteOrder            uint16      // 字节序（0xFFFE为小端序）
	SectorShift          uint16      // 扇区大小（512=0x0009, 4096=0x000C）
	MiniSectorShift      uint16      // 迷你扇区大小（固定64字节 = 0x0006）  @offset = 0x20
	Reserved             [6]byte     // 保留字段
	DirectorySectorCnt   uint32      // 目录扇区数量,MajorVersion=3时为0
	FATSectorCnt         uint32      // FAT表扇区数量
	DirectoryStart       uint32      // 目录起始扇区ID                     @offset = 0x30
	TransactionSignature uint32      // 事务签名（MajorVersion=4时使用）
	MiniStreamCutoffSize uint32      // 迷你流截断大小（MajorVersion=4时使用)
	MiniFATStart         uint32      // 迷你FAT起始扇区ID
	MiniFATSectorCnt     uint32      // 迷你FAT扇区数量
	DiFATSectorStart     uint32      // DIFAT起始扇区ID
	DIFATSectorCnt       uint32      // DIFAT扇区数量
	DiFAT                [109]uint32 // DIFAT扇区ID数组（每个4字节，共109个条目）
}

// 目录项结构 (128字节)
type DirectoryEntry struct {
	Name           [64]byte // UTF-16名称
	NameLen        uint16   // 实际名称长度
	ObjectType     uint8    // 类型：0x0(unknown) 0x01(存储) 0x02(流) 0x05(根存储)
	ColorFlag      uint8    // 颜色标志（0x00=红色, 0x01=黑色）
	LeftSiblingID  uint32   // 左兄弟项ID
	RightSiblingID uint32   // 右兄弟项ID
	ChildID        uint32   // 子项ID
	CLSID          [16]byte // CLSID（保留字段）
	StateBits      uint32   // 状态位（0x00000001=已分配, 0x00000002=已删除）
	CreationTime   int64    // 创建时间（自1601年1月1日起的100纳秒间隔）
	ModifiedTime   int64    // 修改时间（自1601年1月1日起的100纳秒间隔）
	StartSectorID  uint32   // 流起始扇区ID [1,8](@ref)
	StreamSize     uint64   // 流大小
}

type PDirectoryEntry struct {
	Name  string
	Type  uint8
	Entry *DirectoryEntry
}

type DocParse struct {
	File *os.File // 文件句柄

	/*文件头 */
	FileHeader *FileHeader

	/* 目录项 */
	DirEntry []*PDirectoryEntry
	FIB      *fib.Fib // 存储解析后的FIB数据

	/* DIFAT */
	DIFAT   []uint32 // DIFAT扇区ID列表
	FAT     []uint32 //uint32数组，每个元素表示一个扇区ID
	MiniFAT []uint32

	WordDocumentStream []byte

	SectorSize int
	IsMiniFAT  bool

	Table1SectorStartID uint32 // 1Table stream起始ID
	Table1SectorSize    uint64 // 1Table stream大小
	Table0SectorStartID uint32 // 0Table stream起始ID
	Table0SectorSize    uint64 // 0Table stream大小
	MainCharactorNum    uint32 // 主要字符数
	CLXOffset           uint32 // CLX偏移量
	CLXSize             uint32 // CLX大小
}

type OfficeDocParser struct{}

func decodeText(data []byte, encodingFlag byte) string {
	if encodingFlag == 0x00 { // ANSI编码（GBK中文）
		decoder := simplifiedchinese.GBK.NewDecoder()
		result, _ := decoder.String(string(data))
		return result
	} else { // UTF-16LE
		runes := make([]rune, len(data)/2)
		for i := 0; i < len(runes); i++ {
			runes[i] = rune(binary.LittleEndian.Uint16(data[2*i:]))
		}
		return string(runes)
	}
}

// 解码UTF-16字节流为字符串（支持代理对和字节序处理）
func decodeUTF16(data []byte, byteOrder binary.ByteOrder) string {
	// 1. 字节序检测与BOM处理
	var bomSize int
	if len(data) >= 2 {
		switch {
		case data[0] == 0xFE && data[1] == 0xFF:
			byteOrder = binary.BigEndian // 大端序标识
			bomSize = 2
		case data[0] == 0xFF && data[1] == 0xFE:
			byteOrder = binary.LittleEndian // 小端序标识
			bomSize = 2
		}
	}
	if byteOrder == nil {
		byteOrder = binary.LittleEndian // DOC默认小端序
	}

	// 2. 将字节流转换为uint16序列
	u16s := make([]uint16, (len(data)-bomSize)/2)
	for i := 0; i < len(u16s); i++ {
		u16s[i] = byteOrder.Uint16(data[bomSize+2*i:])
	}

	// 3. 处理UTF-16代理对（4字节字符）
	var runes []rune
	for i := 0; i < len(u16s); {
		switch {
		case utf16.IsSurrogate(int32(u16s[i])):
			if i+1 < len(u16s) {
				// 解码代理对（如中文/emoji）
				r := utf16.DecodeRune(rune(u16s[i]), rune(u16s[i+1]))
				runes = append(runes, r)
				i += 2 // 跳过已处理的代理对
			} else {
				// 代理对不完整
				runes = append(runes, utf8.RuneError)
				i++
			}
		default:
			// 基本平面字符（2字节）
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

// 查找RootEntry流
func (e *PDirectoryEntry) CheckRootEntry() bool {
	return e.Type == 0x05
}

func (e *PDirectoryEntry) CheckTextStream() bool {
	// 查找主文本流（WordDocument）
	return e.Type == 0x02 && strings.Contains(e.Name, "WordDocument")
}

func (e *PDirectoryEntry) CheckTable0Straem() bool {
	return e.Type == 0x02 && strings.Contains(e.Name, "0Table")
}

func (e *PDirectoryEntry) CheckTable1Straem() bool {
	return e.Type == 0x02 && strings.Contains(e.Name, "1Table")
}

func (e *PDirectoryEntry) CheckTableStream(fibBase *fib.FibBase) bool {
	// 根据FIB中的fWhichTblStm属性确定Table流名称
	var tableName string
	if (fibBase.Flags & 0x0200) != 0 {
		tableName = "1Table"
	} else {
		tableName = "0Table"
	}
	return e.Type == 0x02 && e.Name == tableName
}

func (e *DirectoryEntry) isMiniStream() bool {
	return e.StreamSize <= 4096
}

func NewDocParse(fn string) (*DocParse, error) {
	file, err := os.Open(fn)
	if err != nil {
		return nil, fmt.Errorf("文件 %s 打开失败: %w", fn, err)
	}
	return &DocParse{
		File:               file,
		FileHeader:         &FileHeader{},
		DirEntry:           make([]*PDirectoryEntry, 0),
		FAT:                make([]uint32, 0),
		DIFAT:              make([]uint32, 0),
		MiniFAT:            make([]uint32, 0),
		WordDocumentStream: make([]byte, 0),
		IsMiniFAT:          false,
	}, nil
}

func (d *DocParse) Close() {
	if d.File != nil {
		d.File.Close()
		d.File = nil
	}
}

func (d *DocParse) ParseHeader() error {
	file := d.File
	header := &FileHeader{}
	if err := binary.Read(file, binary.LittleEndian, header); err != nil {
		return err
	}

	// 验证签名 (偏移0x0000)
	if hex.EncodeToString(header.Signature[:]) != DocSignature {
		return errors.New("无效的OLE签名")
	}

	header.Printf()
	d.SectorSize = 1 << header.SectorShift
	d.FileHeader = header
	return nil
}

func (d *DocParse) GetWordDocumentStream(e *PDirectoryEntry) error {
	var textBuilder bytes.Buffer

	entry := e.Entry
	currentSector := entry.StartSectorID

	logger.Logger.Printf("开始提取文本流，扇区大小：%d, 起始扇区: %d, stream大小: %d\n", d.SectorSize, currentSector, entry.StreamSize)
	// 遍历FAT扇区链
	var pos uint64
	for currentSector != 0xFFFFFFFE { // 0xFFFFFFFE表示链结束
		if pos >= entry.StreamSize {
			break
		}
		// 计算扇区物理位置：文件头后偏移 = 512 + 扇区ID * 扇区大小
		sectorPos := int64(DocHeaderOffset + int(currentSector)*int(d.SectorSize))
		logger.DebugLogger.Printf("文件读取偏移: 0x%x(扇区id:%d), 读取长度：%d, 剩余长度：%d\n", sectorPos, currentSector, pos, entry.StreamSize-pos)

		_, err := d.File.Seek(sectorPos, 0)
		if err != nil {
			return err
		}

		var saved uint64
		if entry.StreamSize-pos >= uint64(d.SectorSize) {
			saved = uint64(d.SectorSize)
		} else {
			saved = entry.StreamSize - pos
		}
		// 读取扇区数据
		sectorData := make([]byte, saved)
		if _, err := d.File.Read(sectorData); err != nil {
			return err
		}

		textBuilder.Write(sectorData)
		pos += saved
		currentSector = d.FAT[currentSector] // 获取下一扇区
	}

	d.WordDocumentStream = textBuilder.Bytes()
	logger.DebugLogger.Printf("worddocument文本流大小： %d\n", len(d.WordDocumentStream))
	return nil
}

func (d *DocParse) UpdateDirectoryInfo(entry *PDirectoryEntry) error {
	if entry.CheckTextStream() {
		if err := d.GetWordDocumentStream(entry); err != nil {
			return err
		}
	} else if entry.CheckRootEntry() {
		// 用于miinfat的查找，暂时不处理
	} else if entry.CheckTable1Straem() {
		d.Table1SectorStartID = entry.Entry.StartSectorID
		d.Table1SectorSize = entry.Entry.StreamSize
		logger.Logger.Printf("Table1 Stream信息: 起始扇区ID: %d, 大小: %d\n", d.Table1SectorStartID, d.Table1SectorSize)
	} else if entry.CheckTable0Straem() {
		d.Table0SectorStartID = entry.Entry.StartSectorID
		d.Table0SectorSize = entry.Entry.StreamSize
		logger.Logger.Printf("Table0 Stream信息: 起始扇区ID: %d, 大小: %d\n", d.Table0SectorStartID, d.Table0SectorSize)
	}

	return nil
}

func (d *DocParse) GetDirEntries() error {
	header := d.FileHeader
	file := d.File

	dirSectorPos := DocHeaderOffset + int64(header.DirectoryStart)*int64(d.SectorSize)
	logger.Logger.Printf("扇区大小：%d, 扇区数量: %d, 开始扇区: 0x%x, 目录扇区起始偏移: 0x%x\n",
		int64(d.SectorSize), header.DirectorySectorCnt, header.DirectoryStart, dirSectorPos)

	_, err := file.Seek(dirSectorPos, 0)
	if err != nil {
		return err
	}

	direntryCount := 0
	if header.MajorVersion == 3 {
		direntryCount = d.SectorSize / 128
	} else {
		direntryCount = int(header.DirectorySectorCnt+1) * (d.SectorSize / 128)
	}

	for i := 0; i < direntryCount; i++ {
		entry := &DirectoryEntry{}
		if err := binary.Read(file, binary.LittleEndian, entry); err != nil {
			break
		}
		if entry.NameLen > 64 {
			logger.Logger.Printf("目录项名称长度超过64字节")
			return nil
		}

		name := decodeUTF16(entry.Name[:entry.NameLen], binary.LittleEndian)
		pd := &PDirectoryEntry{
			Name:  name,
			Type:  entry.ObjectType,
			Entry: entry,
		}
		d.DirEntry = append(d.DirEntry, pd)

		d.UpdateDirectoryInfo(pd)

		logger.Logger.Printf("目录项名称: %s, 长度： %d, 类型: %d, 起始扇区: %d, 大小: %d\n",
			name, entry.NameLen, entry.ObjectType, entry.StartSectorID, entry.StreamSize)
	}

	if len(d.DirEntry) == 0 {
		return errors.New("no directory entry found")
	}
	return nil
}

func (d *DocParse) GetRootEntrySectorStartID() (uint32, bool) {
	for _, entry := range d.DirEntry {
		if entry.CheckRootEntry() {
			return entry.Entry.StartSectorID, true
		}
	}
	return uint32(0), false
}

// 也就是解析FIB
func (d *DocParse) ParseWordDocument() error {
	if len(d.WordDocumentStream) == 0 {
		return fmt.Errorf("no worddocument found\n")
	}

	// 解析FIB文件格式
	fib, err := fib.ParseFIB(d.WordDocumentStream)
	if err != nil {
		return fmt.Errorf("解析FIB文件失败: %w\n", err)
	}

	d.FIB = fib

	// 验证CLX偏移是否有效
	if d.FIB.FcClx == 0 || d.FIB.LcbClx == 0 {
		return fmt.Errorf("未找到有效的CLX偏移信息")
	}
	return nil
}

func (d *DocParse) ParseFibClx() ([]byte, error) {
	var tableOffset uint32
	var tableSize uint64
	tableOffset = DocHeaderOffset + d.Table0SectorStartID*uint32(d.SectorSize)
	tableSize = d.Table0SectorSize
	if d.FIB.Base != nil && d.FIB.Base.Flags&0x0200 != 0 {
		tableOffset = DocHeaderOffset + d.Table1SectorStartID*uint32(d.SectorSize)
		tableSize = d.Table1SectorSize
	}

	logger.DebugLogger.Printf("flag: %v, tableOffset: 0x%x, tableSize: 0x%x\n",
		d.FIB.Base.Flags&0x0200, tableOffset, tableSize)
	return d.FIB.ParseFibClx(d.File, d.WordDocumentStream, tableOffset, tableSize)
}

// 定位
func (d *DocParse) ExtractText() ([]byte, error) {
	return d.ParseFibClx()
}

func (d *DocParse) ExtractEntry(entry *DirectoryEntry, sectorSize uint64, isMini bool) ([]byte, error) {
	var textBuilder bytes.Buffer
	currentSector := entry.StartSectorID

	logger.Logger.Printf("开始提取文本流，起始扇区(%d): %d, 大小: %d\n", sectorSize, currentSector, entry.StreamSize)
	// 遍历FAT扇区链
	var pos uint64
	for currentSector != 0xFFFFFFFE { // 0xFFFFFFFE表示链结束
		if pos >= entry.StreamSize {
			break
		}
		// 计算扇区物理位置：文件头后偏移 = 512 + 扇区ID * 扇区大小
		sectorPos := int64(512 + int(currentSector)*int(sectorSize))
		logger.DebugLogger.Printf("文件读取偏移: 0x%x, 读取长度：%d, 剩余长度：%d\n", sectorPos, pos, entry.StreamSize-pos)
		_, err := d.File.Seek(sectorPos, 0)
		if err != nil {
			return textBuilder.Bytes(), err
		}

		var saved uint64
		if entry.StreamSize-pos >= uint64(sectorSize) {
			saved = sectorSize
		} else {
			saved = entry.StreamSize - pos
		}
		// 读取扇区数据
		sectorData := make([]byte, saved)
		if _, err := d.File.Read(sectorData); err != nil {
			return textBuilder.Bytes(), err
		}

		textBuilder.Write(sectorData)

		//text := decodeText(sectorData, 1)
		//textBuilder.WriteString(text)
		//fmt.Printf("记录内容: %s\n", text)
		pos += saved
		//fmt.Printf("读取记录类型: 0x%04X, 大小: %d, 当前偏移: %d\n", recordType, recordSize, pos)
		currentSector = d.FAT[currentSector] // 获取下一扇区
	}
	return textBuilder.Bytes(), nil
}

func (d *DocParse) LoadFAT() error {
	file := d.File
	fat := make([]uint32, 0)
	entriesPerSector := d.SectorSize / 4 // 每个扇区的FAT条目数

	// 使用DIFAT中的扇区ID读取所有FAT扇区
	for _, fatSectorID := range d.DIFAT {
		if fatSectorID == 0xFFFFFFFF {
			continue // 跳过空条目
		}
		sectorPos := int64(DocHeaderOffset) + int64(fatSectorID)*int64(d.SectorSize)
		_, err := file.Seek(sectorPos, 0)
		if err != nil {
			return err
		}
		// 读取当前FAT扇区的所有条目
		entries := make([]uint32, entriesPerSector)
		if err := binary.Read(file, binary.LittleEndian, &entries); err != nil {
			return err
		}
		fat = append(fat, entries...)
	}

	d.FAT = fat
	logger.DebugLogger.Printf("FAT扇区ID[%d]: %v\n", len(fat), fat)
	return nil
}

func (d *DocParse) LoadMiniFAT() error {
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

func (d *DocParse) LoadDIFAT() error {
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
		sectorPos := DocHeaderOffset + int64(currentSector)*int64(d.SectorSize)
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

func (d *DocParse) TraverseFAT(startSector uint32) ([]uint32, error) {
	var chain []uint32
	current := startSector

	for current != 0xFFFFFFFE { // 0xFFFFFFFE表示链结束
		if int(current) >= len(d.FAT) {
			return nil, fmt.Errorf("无效的FAT索引%d", current)
		}
		chain = append(chain, current)
		current = d.FAT[current] // 获取下一扇区
	}
	return chain, nil
}

func (d *DocParse) TraverseMiniFAT(startSector uint32) ([]uint32, error) {
	var chain []uint32
	current := startSector

	for current != 0xFFFFFFFE {
		if int(current) >= len(d.MiniFAT) {
			return nil, fmt.Errorf("无效的MiniFAT索引%d", current)
		}
		chain = append(chain, current)
		current = d.MiniFAT[current]
	}
	return chain, nil
}

/*
// 提取doc文档中的文本内容
func ExtractDoc(fn string) ([]byte, error) {
	docparser, err := NewDocParse(fn)
	if err != nil {
		fmt.Printf("创建DocParse实例失败: %v\n", err)
		return []byte{}, err
	}
	defer docparser.Close()

	// 1. 解析文件头
	if err = docparser.ParseHeader(); err != nil {
		fmt.Printf("解析文件头失败: %v\n", err)
		return []byte{}, err
	}

	// 2. 解析difat表
	if err = docparser.LoadDIFAT(); err != nil {
		fmt.Printf("加载DIFAT表失败: %v", err)
		return []byte{}, err
	}

	// 3. 加载FAT表
	if err = docparser.LoadFAT(); err != nil {
		fmt.Printf("加载FAT表失败: %v", err)
		return []byte{}, err
	}

	if err = docparser.LoadMiniFAT(); err != nil {
		fmt.Printf("加载MiniFAT表失败: %v", err)
		return []byte{}, err
	}

	if err = docparser.GetDirEntries(); err != nil {
		fmt.Printf("获取目录项失败: %v\n", err)
		return []byte{}, err
	}

	if err = docparser.ParseWordDocument(); err != nil {
		fmt.Printf("解析WordDocument失败: %v", err)
		return []byte{}, err
	}

	content, err := docparser.ExtractText()
	if err != nil {
		fmt.Printf("提取文本内容失败: %v", err)
		return []byte{}, err
	}

	return content, err
}
*/

func (p *OfficeDocParser) Parse(filePath string) ([]byte, error) {
	docparser, err := NewDocParse(filePath)
	if err != nil {
		return []byte{}, fmt.Errorf("创建DocParse实例失败: %w\n", err)
	}
	defer docparser.Close()

	// 1. 解析文件头
	if err = docparser.ParseHeader(); err != nil {
		return []byte{}, fmt.Errorf("解析文件头失败: %w\n", err)
	}

	// 2. 解析difat表
	if err = docparser.LoadDIFAT(); err != nil {
		return []byte{}, fmt.Errorf("加载DIFAT表失败: %w\n", err)
	}

	// 3. 加载FAT表
	if err = docparser.LoadFAT(); err != nil {
		return []byte{}, fmt.Errorf("加载FAT表失败: %w\n", err)
	}

	if err = docparser.LoadMiniFAT(); err != nil {
		return []byte{}, fmt.Errorf("加载MiniFAT表失败: %w\n", err)
	}

	if err = docparser.GetDirEntries(); err != nil {
		return []byte{}, fmt.Errorf("获取目录项失败: %w\n", err)
	}

	if err = docparser.ParseWordDocument(); err != nil {
		return []byte{}, fmt.Errorf("解析WordDocument失败: %w\n", err)
	}

	content, err := docparser.ExtractText()
	if err != nil {
		return []byte{}, fmt.Errorf("提取文本内容失败: %w\n", err)
	}

	return content, err
}

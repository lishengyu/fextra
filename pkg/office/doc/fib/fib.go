package fib

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"os"
	"unsafe"

	"fextra/pkg/logger"
	"fextra/pkg/office/doc/fib/clx"
)

type FibBase struct {
	// 0x000-0x001: 文件标识
	WIdent uint16 // 必须是0xA5EC(word)

	/*
		1. Check the value of FIB.cswNew.
		2. If the value is 0, nFib is specified by FibBase.nFib.
		3. Otherwise, the value is not 0 and nFib is specified by FibRgCswNew.nFibNew.
	*/
	NFib uint16 // 文件格式的版本号  0x006A=Word6, 0x00C1=Word97

	Unused uint16 // 保留

	// 0x006-0x007: 语言标识
	Language uint16 // lid

	PnNext uint16 //文本或图形的偏移

	// 0x008-0x009: 文档类型标记
	Flags uint16 // 位掩码定义 (按位从低到高):
	// 0x0001 (bit 0) - fDot: 文档模板标志
	// 0x0002 (bit 1) - fGlsy: 仅包含自动图文集项
	// 0x0004 (bit 2) - fComplex: 上次保存为增量保存
	// 0x0008 (bit 3) - fHasPic: 文档包含图片
	// 0x00F0 (bits 4-7) - cQuckSaves: 快速保存次数
	// 0x0100 (bit 8) - fEncrypted: 文档已加密或混淆
	// 0x0200 (bit 9) - fWhichTblStm: 表流选择标志(1=1Table, 0=0Table)
	// 0x0400 (bit10) - fReadOnlyRecommended: 建议只读模式
	// 0x0800 (bit11) - fWriteReservation: 有写保护
	// 0x1000 (bit12) - fExtChar: 必须为1
	// 0x2000 (bit13) - fLoadOverride: 覆盖语言和字体信息
	// 0x4000 (bit14) - fFarEast: 创建文档的应用语言为东亚语言
	// 0x8000 (bit15) - fObfuscated: 文档使用XOR混淆

	NFlibBack uint16 // must be one of 0x00BF or 0x00C1

	IKey uint32

	Envr uint8 // must be 0

	Reserved1 uint8

	Reserved3 uint16
	Reserved4 uint16
	Reserved5 uint32
	Reserved6 uint32
}

type FibRgCswNew struct {
	NFibNew      uint16  // one of 0x00D9 0x0101 0x010C 0x0112
	RgCswNewData []uint8 // depend on NfibNew
	/*
		0x00D9 fibRgCswNewData2000 (2 bytes)
		0x0101 fibRgCswNewData2000 (2 bytes)
		0x010C fibRgCswNewData2000 (2 bytes)
		0x0112 fibRgCswNewData2007 (8 bytes)
	*/
}

type FibRgW97 struct {
	FibRgw [28]uint8
}

type FibRgLw97 struct {
	FibRgLw [88]uint8
}

// ================================================
const (
	CcpTextIndex = 3 //主文档中的字符数量
)
const (
	FcClxIndex  = 66 // clx offset，在FibRgFclcb97中的索引
	LcbClxIndex = 67 // clx大小，单位bytes
)

// 查找clx数据结构  ==>   查找prc数据结构

// 接下来都是FibRgFclcb结构，需要根据nlib来确认是什么结构
type FibRgFclcb97 struct {
}

//================================================

type Fib struct {
	Reader         *bytes.Reader
	Base           *FibBase
	Csw            uint16    // must be 0x000e
	FibRgw         FibRgW97  // Csw * FibRgw(28 bytes)
	Cslw           uint16    // 0x0016
	FibRgLw        FibRgLw97 // Cslw * FibRgLw(88 bytes)
	CbRgFcLcb      uint16    // depend on nFib
	FibRgFcLcbBlob []uint8   // depend on nFib  类似于union 取不同的数据类型
	/*
		0x00C1 fibRgFcLcb97
		0x00D9 fibRgFcLcb2000
		0x0101 fibRgFcLcb2002
		0x010C fibRgFcLcb2003
		0x0112 fibRgFcLcb2007
	*/
	CswNew      uint16 // depend on nFib
	FibRgCswNew []FibRgCswNew

	CcpText uint32 // 主文本字符数量
	FcClx   uint32 // Table Stream中文本偏移位置
	LcbClx  uint32 // Table Stream中文本大小
}

func (fb *FibBase) Printf() {
	logger.DebugLogger.Printf("FibBase:\n")
	logger.DebugLogger.Printf("WIdent: 0x%x\n", fb.WIdent)
	logger.DebugLogger.Printf("NFib: 0x%x\n", fb.NFib)
	logger.DebugLogger.Printf("Unused: 0x%x\n", fb.Unused)
	logger.DebugLogger.Printf("Language: 0x%x\n", fb.Language)
	logger.DebugLogger.Printf("PnNext: 0x%x\n", fb.PnNext)
	logger.DebugLogger.Printf("Flags: 0x%x\n", fb.Flags)
	logger.DebugLogger.Printf("NFlibBack: 0x%x\n", fb.NFlibBack)
	logger.DebugLogger.Printf("IKey: 0x%x\n", fb.IKey)
	logger.DebugLogger.Printf("Envr: 0x%x\n", fb.Envr)
}

func (f *Fib) parseFibCsw() error {
	var cswCnt uint16
	if err := binary.Read(f.Reader, binary.LittleEndian, &cswCnt); err != nil {
		return err
	}

	if cswCnt != 0x000E {
		return fmt.Errorf("invalid cswCnt: %d\n", cswCnt)
	}

	buf := make([]byte, 2*cswCnt)
	if _, err := io.ReadFull(f.Reader, buf); err != nil {
		return err
	}
	logger.DebugLogger.Printf("csw count: %d\n", cswCnt)
	csw := make([]uint16, cswCnt)
	for i := range csw {
		csw[i] = binary.LittleEndian.Uint16(buf[2*i:])
		logger.DebugLogger.Printf("%d(0x%x) \n", i, csw[i])
	}
	logger.DebugLogger.Printf("\n====> end\n")
	return nil
}

func (f *Fib) parseFibCslw() error {
	var cslwCnt uint16

	if err := binary.Read(f.Reader, binary.LittleEndian, &cslwCnt); err != nil {
		return err
	}
	if cslwCnt != 0x0016 {
		return fmt.Errorf("invalid cslwCnt: %d\n", cslwCnt)
	}

	logger.DebugLogger.Printf("cslw count: %d\n", cslwCnt)
	buf := make([]byte, 4*cslwCnt)
	if _, err := io.ReadFull(f.Reader, buf); err != nil {
		return err
	}
	cslw := make([]uint32, cslwCnt)
	for i := range cslw {
		cslw[i] = binary.LittleEndian.Uint32(buf[4*i:])
		logger.DebugLogger.Printf("%d(0x%x) \n", i, cslw[i])
		if i == CcpTextIndex {
			f.CcpText = cslw[i]
		}
	}
	logger.DebugLogger.Printf("\n====> end\n")
	return nil
}

// 临时存放，确认解析逻辑是否正确
var (
	tempOffset int
	TempFcClx  uint32
	TempLcbClx uint32
)

func (f *Fib) parseFibFclcb(nfib uint16) error {
	var fclcbCnt uint16

	if err := binary.Read(f.Reader, binary.LittleEndian, &fclcbCnt); err != nil {
		return err
	}

	if fclcbCnt != 0x005D && fclcbCnt != 0x006C && fclcbCnt != 0x0088 && fclcbCnt != 0x00A4 && fclcbCnt != 0x00B7 {
		return fmt.Errorf("invalid fclcb: %d\n", fclcbCnt)
	}

	logger.DebugLogger.Printf("cslw count: %d\n", fclcbCnt)
	buf := make([]byte, 8*fclcbCnt)
	if _, err := io.ReadFull(f.Reader, buf); err != nil {
		return err
	}
	fclcb := make([]uint32, fclcbCnt*2)
	for i := range fclcb {
		fclcb[i] = binary.LittleEndian.Uint32(buf[4*i:])
		logger.DebugLogger.Printf("%d(0x%x)\n", i, fclcb[i])
	}

	if len(fclcb) >= FcClxIndex && len(fclcb) >= LcbClxIndex {
		f.FcClx = fclcb[FcClxIndex]
		f.LcbClx = fclcb[LcbClxIndex]
		logger.Logger.Printf("提取CLX偏移: 0x%x, 大小: %d字节\n", f.FcClx, f.LcbClx)
	}
	logger.DebugLogger.Printf("\n====> end\n")
	return nil
}

func (f *Fib) parseFibCswNew() error {
	var cswNewCnt uint16

	if err := binary.Read(f.Reader, binary.LittleEndian, &cswNewCnt); err != nil {
		return err
	}

	if cswNewCnt != 0x005D && cswNewCnt != 0x006C && cswNewCnt != 0x0088 && cswNewCnt != 0x00A4 && cswNewCnt != 0x00B7 {
		return fmt.Errorf("invalid cswNew: %d\n", cswNewCnt)
	}

	cswNew := make([]uint16, cswNewCnt)
	if err := binary.Read(f.Reader, binary.LittleEndian, &cswNew); err != nil {
		return err
	}
	tempOffset = tempOffset + 2 + int(unsafe.Sizeof(cswNew))
	logger.DebugLogger.Printf("totaol %d cswNew offset += %x hex %x\n", tempOffset, len(cswNew), cswNew[:])
	return nil
}

func (f *Fib) ParseFibBase() error {
	// fibbase固定字节
	fibBase := &FibBase{}
	//fibBase := make([]uint8, 32)
	if err := binary.Read(f.Reader, binary.LittleEndian, fibBase); err != nil {
		return err
	}
	f.Base = fibBase
	fibBase.Printf()
	return nil
}

func (f *Fib) ParseFibClx(r *os.File, wd []byte, offset uint32, size uint64) ([]byte, error) {
	clxOffset := offset + f.FcClx
	logger.DebugLogger.Printf("clxoffset: 0x%x\n", clxOffset)
	_, err := r.Seek(int64(clxOffset), 0)
	if err != nil {
		return []byte{}, err
	}

	buf := make([]byte, f.LcbClx)
	if _, err = io.ReadFull(r, buf); err != nil {
		return []byte{}, err
	}

	// 此处偏移已经定位到clx，直接按照clx进行解析
	clxData, err := clx.ParseClx(buf)
	if err != nil {
		return []byte{}, err
	}
	pcdt := &clxData.Pcdt

	// 提取pcdt中的纯文本内容
	var textBuilder bytes.Buffer
	acp := pcdt.PlcPcd.ACP
	apcd := pcdt.PlcPcd.APcd

	logger.DebugLogger.Printf("acp: %v\n", acp)
	for _, v := range apcd {
		logger.DebugLogger.Printf("FcCompressed: 0x%x, Flags: 0x%x, Prm: 0x%x, 0x%x, %v\n", v.FcCompressed, v.Flags, v.Prm, v.Fc(), v.IsCompressed())
	}

	for i := 0; i < len(apcd); i++ {
		startCp := acp[i]
		endCp := acp[i+1]
		length := endCp - startCp

		if length == 0 {
			continue
		}

		logger.DebugLogger.Printf("startcp: %d, endcp: %d, length: %d, charnum: %d, data len: %d\n",
			startCp, endCp, length, size, len(buf))

		segment, err := pcdt.GetText(startCp, f.CcpText, wd)
		if err != nil {
			return []byte{}, fmt.Errorf("提取文本片段失败(索引%d): %w", i, err)
		}
		logger.DebugLogger.Printf("content[%d]:\n%s\n", len(segment), segment)
		textBuilder.WriteString(segment)
	}

	return textBuilder.Bytes(), nil
}

func NewFib(data []byte) *Fib {
	return &Fib{
		Reader: bytes.NewReader(data),
	}
}

func ParseFIB(data []byte) (*Fib, error) {
	nf := NewFib(data)

	if err := nf.ParseFibBase(); err != nil {
		return nf, err
	}

	if err := nf.parseFibCsw(); err != nil {
		return nf, err
	}

	if err := nf.parseFibCslw(); err != nil {
		return nf, err
	}

	if err := nf.parseFibFclcb(nf.Base.NFib); err != nil {
		return nf, err
	}

	/*
		if err := nf.parseFibCswNew(); err != nil {
			return nf, err
		}
	*/

	return nf, nil
}

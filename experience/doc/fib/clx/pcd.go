package clx

import (
	"encoding/binary"
	"errors"
	"fextra/pkg/logger"
	"fmt"
	"math"
	"sort"
	"unicode/utf16"

	"golang.org/x/text/encoding/simplifiedchinese"
	"golang.org/x/text/transform"
)

// Pcd结构定义 (8字节)
// 参考: 2.9.177 Pcd规范
// 结构标识常量
const (
	PrcClxtIdentifier  = 0x01 // PRC结构标识 (2.9.209 Prc规范)
	PcdtClxtIdentifier = 0x02 // PCDT结构标识 (2.9.178 Pcdt规范)
)

// Pcd标志位掩码
const (
	PcdFlagANoParaLast = 1 << 0 // 0x0001: 无段落标记
	PcdFlagBReserved   = 1 << 1 // 0x0002: 保留位
	PcdFlagCDirty      = 1 << 2 // 0x0004: 脏标记
	PcdFlagFR2Mask     = 0xFFF8 // 0xFFF8: FR2保留字段
)

// Pcd结构定义 (8字节)
// 参考: 2.9.177 Pcd规范
type Pcd struct {
	Flags        uint16 // 标志位: A(1), B(1), C(1), FR2(13)
	FcCompressed uint32 // 32位结构 (2.9.73):
	// - 位0-29: fc (文本流偏移)
	// - 位30: A (FCompressed标志)
	// - 位31: B (保留位，必须为0)
	Prm uint16 // 段落/字符属性 (Prm结构)
}

// Fc 返回文本流偏移值 (30位)
func (p *Pcd) Fc() uint32 {
	return p.FcCompressed & 0x3FFFFFFF
}

// IsCompressed 返回文本是否压缩 (A标志位)
func (p *Pcd) IsCompressed() bool {
	return (p.FcCompressed & 0x40000000) != 0
}

// ValidateReservedBit 验证保留位必须为0
func (p *Pcd) ValidateReservedBit() error {
	if (p.FcCompressed & 0x80000000) != 0 {
		return fmt.Errorf("FcCompressed保留位(B)必须为0，实际值: %08x", p.FcCompressed)
	}
	return nil
}

// PlcPcd结构定义
// 参考: 2.8.35 PlcPcd规范
type PlcPcd struct {
	ACP  []uint32 // 文本范围起始点数组
	APcd []Pcd    // Pcd结构数组 (每个8字节)
}

// Pcdt结构定义
// 参考: 2.9.178 Pcdt规范
type Pcdt struct {
	Clxt   byte   // 1字节标识，必须为0x02
	Lcb    uint32 // PlcPcd结构大小 (字节)
	PlcPcd PlcPcd // PlcPcd结构
}

// GetText 根据字符位置(cp)从WordDocument流提取文本
// 参考: 2.4.1 Retrieving Text规范
func (pcdt *Pcdt) GetText(cp uint32, length uint32, wordDocStream []byte) (string, error) {
	// 步骤1: 验证参数有效性
	if length == 0 {
		return "", errors.New("提取长度(length)不能为0")
	}
	if length > math.MaxUint32/2 {
		return "", errors.New("提取长度超出最大限制")
	}

	// 步骤2: 在ACP数组中查找最大的<= cp的索引
	acp := pcdt.PlcPcd.ACP
	apcd := pcdt.PlcPcd.APcd

	// ACP数组必须比APcd数组多一个元素
	if len(acp) != len(apcd)+1 {
		return "", fmt.Errorf("ACP数组长度(%d)必须比APcd数组长度(%d)多1，可能是CLX结构损坏", len(acp), len(apcd))
	}

	// 验证ACP数组是否按升序排列
	for j := 1; j < len(acp); j++ {
		if acp[j] < acp[j-1] {
			return "", fmt.Errorf("ACP数组未按升序排列，在索引%d处发现异常", j)
		}
	}

	// 检查cp是否超出有效范围
	if len(acp) == 0 || cp > acp[len(acp)-1] {
		return "", fmt.Errorf("字符位置%d超出有效范围(最大%d)", cp, acp[len(acp)-1])
	}

	// 使用二分查找找到索引i
	i := sort.Search(len(acp), func(j int) bool { return acp[j] > cp }) - 1
	if i < 0 || i >= len(apcd) {
		return "", fmt.Errorf("找不到对应的Pcd条目，cp=%d", cp)
	}

	// 步骤3: 获取对应的Pcd条目
	pcd := apcd[i]
	charOffset := cp - acp[i]

	// 验证提取长度不超出当前Pcd条目范围
	maxCharsInEntry := acp[i+1] - acp[i]
	if charOffset+length > maxCharsInEntry {
		return "", fmt.Errorf("提取长度%d超出当前Pcd条目范围(最大%d字符)", length, maxCharsInEntry)
	}

	// 步骤4: 计算文本流偏移
	fc := pcd.Fc()
	var textOffset uint32

	if pcd.IsCompressed() {
		// 压缩文本: 8-bit ANSI
		textOffset = fc/2 + charOffset
	} else {
		// 未压缩文本: 16-bit Unicode
		textOffset = fc + 2*charOffset
	}

	logger.Logger.Printf("pcd text offset: %d\n", textOffset)
	// 步骤5: 验证偏移有效性
	if textOffset >= uint32(len(wordDocStream)) {
		return "", fmt.Errorf("文本偏移%d超出WordDocument流长度%d", textOffset, len(wordDocStream))
	}

	// 步骤6: 提取并解码文本
	if pcd.IsCompressed() {
		// 压缩文本: 8-bit ANSI (通常为GBK编码)
		byteLength := length
		if textOffset+byteLength > uint32(len(wordDocStream)) {
			return "", fmt.Errorf("压缩文本数据不足(需要%d字节, 实际剩余%d字节)", byteLength, len(wordDocStream)-int(textOffset))
		}
		// 使用GBK解码ANSI文本
		decoder := simplifiedchinese.GBK.NewDecoder()
		result, _, err := transform.Bytes(decoder, wordDocStream[textOffset:textOffset+byteLength])
		if err != nil {
			// 解码失败时返回原始字节的字符串表示
			return string(wordDocStream[textOffset : textOffset+byteLength]), fmt.Errorf("GBK解码失败: %w", err)
		}
		return string(result), nil
	} else {
		// 未压缩文本: 16-bit Unicode (UTF-16LE)
		byteLength := length * 2
		if textOffset+byteLength > uint32(len(wordDocStream)) {
			return "", fmt.Errorf("未压缩文本数据不足(需要%d字节, 实际剩余%d字节)", byteLength, len(wordDocStream)-int(textOffset))
		}
		// 转换字节为uint16切片
		utf16Chars := make([]uint16, length)
		for j := uint32(0); j < length; j++ {
			utf16Chars[j] = binary.LittleEndian.Uint16(wordDocStream[textOffset+j*2:])
		}
		// 解码UTF-16为字符串
		return string(utf16.Decode(utf16Chars)), nil
	}
}

// 解析Pcdt结构
// 参考: 2.9.178 Pcdt规范
func parsePcdt(data []byte) (*Pcdt, error) {
	if len(data) < 5 {
		return nil, errors.New("Pcdt数据不足(至少需要5字节)")
	}

	pcdt := &Pcdt{
		Clxt: data[0],
		Lcb:  binary.LittleEndian.Uint32(data[1:5]),
	}

	// 验证Pcdt标识
	if pcdt.Clxt != PcdtClxtIdentifier {
		return nil, fmt.Errorf("无效Pcdt标识: 0x%x (预期0x%x)", pcdt.Clxt, PcdtClxtIdentifier)
	}

	// 验证Lcb大小
	if pcdt.Lcb == 0 {
		return nil, errors.New("Pcdt.Lcb不能为0")
	}
	if 5+pcdt.Lcb > uint32(len(data)) {
		return nil, fmt.Errorf("Pcdt数据截断 (需要%d字节, 实际%d字节)", 5+pcdt.Lcb, len(data))
	}

	// 解析PlcPcd结构
	plcData := data[5 : 5+pcdt.Lcb]
	acpCount := (len(plcData) + 8) / 12 // (n+1)*4 + n*8 = 12n+4 = Lcb → n=(Lcb-4)/12 → acpCount=n+1=(Lcb+8)/12
	if acpCount == 0 {
		return nil, errors.New("无效的PlcPcd结构: ACP数组为空")
	}

	// 验证PLC数据大小是否匹配公式
	if (acpCount-1)*12+4 != len(plcData) {
		return nil, fmt.Errorf("PlcPcd大小不匹配 (计算大小%d, 实际大小%d)", (acpCount-1)*12+4, len(plcData))
	}

	pcdt.PlcPcd.ACP = make([]uint32, acpCount)

	// 解析aCP数组
	for i := 0; i < acpCount; i++ {
		offset := i * 4
		if offset+4 > len(plcData) {
			return nil, fmt.Errorf("aCP[%d]数据不足", i)
		}
		pcdt.PlcPcd.ACP[i] = binary.LittleEndian.Uint32(plcData[offset:])
	}

	// 解析aPcd数组
	pcdCount := acpCount - 1
	pcdt.PlcPcd.APcd = make([]Pcd, pcdCount)
	for i := 0; i < pcdCount; i++ {
		offset := acpCount*4 + i*8
		if offset+8 > len(plcData) {
			return nil, fmt.Errorf("aPcd[%d]数据不足", i)
		}
		pcdt.PlcPcd.APcd[i] = Pcd{
			Flags:        binary.LittleEndian.Uint16(plcData[offset:]),
			FcCompressed: binary.LittleEndian.Uint32(plcData[offset+2:]),
			Prm:          binary.LittleEndian.Uint16(plcData[offset+6:]),
		}
	}

	return pcdt, nil
}

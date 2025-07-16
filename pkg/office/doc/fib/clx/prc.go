package clx

import (
	"errors"
	"fmt"
)

// RgPrc结构定义
// 参考: 2.9.209 RgPrc规范
type RgPrc struct {
	Clxt byte   // 1字节标识，必须为0x01
	Data []byte // 可变长度的PrcData数据
}

// 解析PRC结构
// 返回解析后的Prc、占用字节数和错误信息
func ParsePrc(data []byte) (RgPrc, int, error) {
	if len(data) < 2 {
		return RgPrc{}, 0, errors.New("PRC数据不足")
	}

	prc := RgPrc{
		Clxt: data[0],
	}

	// 验证Clxt必须为0x01
	if prc.Clxt != PrcClxtIdentifier {
		return RgPrc{}, 0, fmt.Errorf("无效PRC标识: 0x%x (预期0x%x)", prc.Clxt, PrcClxtIdentifier)
	}

	// 解析数据长度 (1字节)
	dataLen := int(data[1])
	dataEnd := 2 + dataLen
	if dataEnd > len(data) {
		return RgPrc{}, 0, fmt.Errorf("PRC数据截断 (需要%d字节, 实际%d字节)", dataEnd, len(data))
	}

	// 复制数据内容
	prc.Data = make([]byte, dataLen)
	copy(prc.Data, data[2:dataEnd])

	return prc, dataEnd, nil
}

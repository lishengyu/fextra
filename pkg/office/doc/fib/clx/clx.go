package clx

import (
	"errors"
	"fextra/pkg/logger"
	"fmt"
)

type Clx struct {
	Prcs []RgPrc // 0,1,or more prcs
	Pcdt Pcdt    // must exist
}

// 解析CLX结构，返回PRC列表和PCDT结构
// 参考: 2.9.209 Prc和2.9.178 Pcdt综合规范
func ParseClx(data []byte) (Clx, error) {
	if len(data) == 0 {
		return Clx{}, errors.New("空CLX数据")
	}

	var prcList []RgPrc
	offset := 0

	// 解析RgPrc数组
	if data[0] == PrcClxtIdentifier {
		prcList = make([]RgPrc, 0)
		for offset < len(data) {
			// 防止无限循环：如果连续100字节没有找到PRC起始或PCDT标识，则认为数据异常
			if offset > 0 && offset%100 == 0 {
				return Clx{}, fmt.Errorf("在偏移%d处未找到有效PRC或PCDT标识", offset)
			}

			if data[0] == PrcClxtIdentifier {
				prc, size, err := ParsePrc(data[offset:])
				if err != nil {
					return Clx{}, fmt.Errorf("解析PRC失败: %w", err)
				}
				prcList = append(prcList, prc)
				offset += size
			} else if data[offset] == PcdtClxtIdentifier {
				// 找到PCDT起始标识，停止PRC解析
				// offset++ // 跳过0x02标识字节
				break
			} else {
				return Clx{}, fmt.Errorf("在偏移%d处未找到有效PRC或PCDT标识", offset)
			}
		}
	} else if data[0] != PcdtClxtIdentifier {
		return Clx{}, fmt.Errorf("在偏移%d处未找到有效PRC或PCDT标识", offset)
	}

	// 验证剩余数据是否足够解析PCDT
	if offset > len(data) {
		return Clx{Prcs: prcList}, errors.New("PCDT数据为空")
	}

	// 解析Pcdt结构
	pcdt, err := parsePcdt(data[offset:])
	if err != nil {
		return Clx{Prcs: prcList}, fmt.Errorf("解析Pcdt失败: %w", err)
	}

	logger.DebugLogger.Printf("pctd: %v\n", pcdt)
	// 验证所有Pcd条目保留位
	for i, pcd := range pcdt.PlcPcd.APcd {
		if err := pcd.ValidateReservedBit(); err != nil {
			return Clx{Prcs: prcList}, fmt.Errorf("Pcd[%d]验证失败: %w", i, err)
		}
	}

	return Clx{
		Prcs: prcList,
		Pcdt: *pcdt,
	}, nil
}

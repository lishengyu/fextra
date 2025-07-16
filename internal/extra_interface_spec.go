package internal

import (
	"path/filepath"
	"strings"
)

// 文件类型常量定义
const (
	FileTypeHTML  = 1
	FileTypeTXT   = 2
	FileTypeXML   = 3
	FileTypeJSON  = 4
	FileTypeCSV   = 5
	FileTypeDOC   = 7
	FileTypeDOCX  = 8
	FileTypeXLS   = 9
	FileTypeXLSX  = 10
	FileTypePPT   = 11
	FileTypePPTX  = 12
	FileTypePDF   = 13
	FileTypeXLSB  = 14
	FileTypeODT   = 15
	FileTypeRTF   = 16
	FileTypeTAR   = 18
	FileTypeGZ    = 19
	FileTypeTARGZ = 20
	FileTypeZIP   = 21
	FileType7Z    = 22
	FileTypeRAR   = 23
	FileTypeBZ2   = 24
	FileTypeJAR   = 25
	FileTypeWAR   = 26
	FileTypeARJ   = 27
	FileTypeLZH   = 28
	FileTypeXZ    = 29
	FileTypeJPEG  = 31
	FileTypePNG   = 32
	FileTypeTIF   = 33
	FileTypeWebP  = 34
	FileTypeWBMP  = 35
	FileTypeVSDX  = 201
	FileTypeVSD   = 202
	FileTypeFPX   = 401
	FileTypePBM   = 402
	FileTypePGM   = 403
	FileTypeBMP   = 404
)

// 定义后缀映射表
var suffixMap = map[string]int{
	"html":   FileTypeHTML,
	"txt":    FileTypeTXT,
	"xml":    FileTypeXML,
	"json":   FileTypeJSON,
	"csv":    FileTypeCSV,
	"doc":    FileTypeDOC,
	"docx":   FileTypeDOCX,
	"xls":    FileTypeXLS,
	"xlsx":   FileTypeXLSX,
	"ppt":    FileTypePPT,
	"pptx":   FileTypePPTX,
	"pdf":    FileTypePDF,
	"xlsb":   FileTypeXLSB,
	"odt":    FileTypeODT,
	"rtf":    FileTypeRTF,
	"vsdx":   FileTypeVSDX,
	"vsd":    FileTypeVSD,
	"tar":    FileTypeTAR,
	"gz":     FileTypeGZ,
	"tar.gz": FileTypeTARGZ,
	"zip":    FileTypeZIP,
	"7z":     FileType7Z,
	"rar":    FileTypeRAR,
	"bz2":    FileTypeBZ2,
	"jar":    FileTypeJAR,
	"war":    FileTypeWAR,
	"arj":    FileTypeARJ,
	"lzh":    FileTypeLZH,
	"xz":     FileTypeXZ,
	"jpeg":   FileTypeJPEG,
	"jpg":    FileTypeJPEG,
	"png":    FileTypePNG,
	"tif":    FileTypeTIF,
	"tiff":   FileTypeTIF,
	"webp":   FileTypeWebP,
	"wbmp":   FileTypeWBMP,
	"fpx":    FileTypeFPX,
	"pbm":    FileTypePBM,
	"pgm":    FileTypePGM,
	"bmp":    FileTypeBMP,
}

// 判断属于哪个大类的其他类型，扩展的其他文件类型
var (
	textOtherSuffixes     = []string{"md", "css", "js", "log", "ini", "py", "go", "java", "c", "cpp", "h", "sh", "bat", "php", "rb"}
	docOtherSuffixes      = []string{"odp", "ods", "pages", "key", "numbers", "wpd"}
	compressOtherSuffixes = []string{"zipx", "tar.bz2", "tar.xz", "rar5", "z"}
	imageOtherSuffixes    = []string{"gif", "ico", "svg", "jpe"}
)

func GetDynamicFileType(filename string) int {
	lowerFilename := strings.ToLower(filename)
	ext := ""

	// 检查复合后缀
	if strings.HasSuffix(lowerFilename, "tar.gz") {
		ext = "tar.gz"
	} else {
		ext = strings.TrimPrefix(filepath.Ext(lowerFilename), ".")
	}

	// 查找后缀对应的FileType
	if t, ok := suffixMap[ext]; ok {
		return t
	} else {
		// 检查是否属于其他文本类
		for _, s := range textOtherSuffixes {
			if ext == s {
				return 6
			}
		}

		// 检查是否属于其他文件类（文档类）
		for _, s := range docOtherSuffixes {
			if ext == s {
				return 17
			}
		}

		// 检查是否属于其他压缩文件类
		for _, s := range compressOtherSuffixes {
			if ext == s {
				return 30
			}
		}

		// 检查是否属于其他图片类
		for _, s := range imageOtherSuffixes {
			if ext == s {
				return 36
			}
		}

		// 都不属于，则设为其他类
		return 114
	}
}

package trie

type TrieNode struct {
	children map[rune]*TrieNode
	isEnd    bool // 标记是否为完整键的结尾（非必须）
}

type PrefixMatcher struct {
	root *TrieNode
}

func NewPrefixMatcher(keys []string) *PrefixMatcher {
	root := &TrieNode{children: make(map[rune]*TrieNode)}
	for _, key := range keys {
		node := root
		for _, ch := range key {
			if node.children[ch] == nil {
				node.children[ch] = &TrieNode{children: make(map[rune]*TrieNode)}
			}
			node = node.children[ch]
		}
		node.isEnd = true
	}
	return &PrefixMatcher{root: root}
}

func (m *PrefixMatcher) HasPrefix(s string) bool {
	node := m.root
	for _, ch := range s {
		if node.children[ch] == nil {
			return false // 字符不匹配，前缀不存在
		}
		node = node.children[ch]
	}
	return true // 成功匹配前缀
}

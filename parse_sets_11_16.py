import json, re

with open('user_text.txt', 'r', encoding='utf-8') as f:
    user_text = f.read()

sets_parsed = {}
curr_set = None

for line in user_text.split('\n'):
    line = line.strip()
    if not line:
        continue
    if line.startswith('SET '):
        curr_set = line.split(' — ')[0].replace('SET', 'Set').strip()
        sets_parsed[curr_set] = []
        continue
    
    # Match Kanji line e.g. "1. 標 (hyou ヒョウ) — signpost"
    m_kanji = re.match(r'^\d+\.\s+([^\s—(]+)\s+.*?—\s*(.*)$', line)
    if m_kanji:
        k_char = m_kanji.group(1)
        k_meaning = m_kanji.group(2)
        if curr_set:
            sets_parsed[curr_set].append((k_char, k_meaning, []))
        continue
    
    # Match Vocab line e.g. "目標 (もくひょう) mokuhyou — goal, target..."
    m_vocab = re.match(r'^([^\s（(]+)\s+[（(]([^）)]+)[）)]\s+([^\s—]+)\s+—\s*(.*)$', line)
    if m_vocab and curr_set and sets_parsed[curr_set]:
        w_kanji = m_vocab.group(1)
        w_hira = m_vocab.group(2)
        w_eng = m_vocab.group(4)
        sets_parsed[curr_set][-1][2].append((w_kanji, w_hira, w_eng))

print('Parsed Sets 11 to 16:')
for s, klist in sets_parsed.items():
    print(f'{s}: {len(klist)} kanji')

with open('sets_11_16.json', 'w', encoding='utf-8') as f:
    json.dump(sets_parsed, f, ensure_ascii=False, indent=2)

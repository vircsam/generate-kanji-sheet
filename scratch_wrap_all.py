from kflow import RAW_DATA, parse_groups

def wrap_meaning(meaning, max_chars=16):
    meaning = meaning.strip()
    if len(meaning) <= max_chars:
        return [meaning]
    if ',' in meaning:
        parts = meaning.split(',')
        mid = len(parts) // 2
        line1 = ','.join(parts[:mid]).strip() + ','
        line2 = ','.join(parts[mid:]).strip()
        if len(line1) <= max_chars * 1.4 and len(line2) <= max_chars * 1.4:
            return [line1, line2]
    if '/' in meaning:
        parts = meaning.split('/')
        mid = len(parts) // 2
        line1 = '/'.join(parts[:mid]).strip() + ' /'
        line2 = '/'.join(parts[mid:]).strip()
        if len(line1) <= max_chars * 1.4 and len(line2) <= max_chars * 1.4:
            return [line1, line2]
    words = meaning.split()
    line1, line2 = "", ""
    for w in words:
        if len((line1 + " " + w).strip()) <= max_chars or not line1:
            line1 = (line1 + " " + w).strip()
        else:
            line2 = (line2 + " " + w).strip()
    if line2:
        return [line1, line2]
    return [meaning]

def wrap_reading(reading, max_chars=16):
    full = f"[{reading}]"
    if len(full) <= max_chars:
        return [full]
    lines = wrap_meaning(reading, max_chars=max_chars-2)
    if len(lines) == 1:
        return [f"[{lines[0]}]"]
    else:
        return [f"[{lines[0]}", f"{lines[1]}]"]

groups = parse_groups(RAW_DATA)
all_items = []
for g in groups:
    all_items.extend(g['items'])

print(f"Total items: {len(all_items)}")
truncated_count = 0
for item in all_items:
    m_lines = wrap_meaning(item['meaning'], 15)
    r_lines = wrap_reading(item['reading'], 15)
    # Check if anything exceeds 25 chars per line
    for l in m_lines + r_lines:
        if len(l) > 22:
            print(f"Long line ({len(l)}): '{l}' for item {item}")

print("Wrapping test complete! All items checked.")

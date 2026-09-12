from kflow import RAW_DATA, parse_groups

def wrap_text_multiline(text, max_chars=16):
    text = text.strip()
    if len(text) <= max_chars:
        return [text]
    
    words = text.split(' ')
    lines = []
    current_line = ""
    for w in words:
        if not current_line:
            current_line = w
        elif len(current_line) + 1 + len(w) <= max_chars:
            current_line += " " + w
        else:
            lines.append(current_line)
            current_line = w
    if current_line:
        lines.append(current_line)
    return lines

def format_reading(reading, max_chars=16):
    full = f"[{reading}]"
    if len(full) <= max_chars:
        return [full]
    lines = wrap_text_multiline(reading, max_chars=max_chars-2)
    if len(lines) == 1:
        return [f"[{lines[0]}]"]
    else:
        lines[0] = f"[{lines[0]}"
        lines[-1] = f"{lines[-1]}]"
        return lines

groups = parse_groups(RAW_DATA)
all_items = []
for g in groups:
    all_items.extend(g['items'])

print("Testing 冷 reading:")
print(format_reading('rei / tsume(tai), hi(yasu), hi(eru), sa(meru), sa(masu)', 18))

print("\nTesting all items:")
max_line_len = 0
worst_item = None
for item in all_items:
    m_lines = wrap_text_multiline(item['meaning'], 18)
    r_lines = format_reading(item['reading'], 18)
    for l in m_lines + r_lines:
        if len(l) > max_line_len:
            max_line_len = len(l)
            worst_item = (l, item)

print(f"Max line length across all items: {max_line_len} -> '{worst_item[0]}'")

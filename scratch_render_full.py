import os
import re
import math
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import subprocess

FONT_PATH = '/System/Library/Fonts/Supplemental/Arial Unicode.ttf'
pdfmetrics.registerFont(TTFont('ArialUnicode', FONT_PATH))

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

# Color Palette
BG_COLOR = colors.HexColor('#FFFFFF')
HEADER_GREEN = colors.HexColor('#1E8449')
HEADER_TEXT_MAIN = colors.HexColor('#0F172A')
HEADER_TEXT_SUB = colors.HexColor('#475569')

RADICAL_BOX_BG = colors.HexColor('#1E293B')
RADICAL_BOX_TXT = colors.HexColor('#FFFFFF')

KANJI_BOX_BG = colors.HexColor('#F8FAFC')
KANJI_BOX_BORDER = colors.HexColor('#475569')
KANJI_CHAR_COLOR = colors.HexColor('#8C2A1E')
MEANING_COLOR = colors.HexColor('#1E293B')
READING_COLOR = colors.HexColor('#4B5563')
LINE_COLOR = colors.HexColor('#94A3B8')

def draw_header_footer(c, page_num, total_pages):
    c.setFillColor(HEADER_GREEN)
    c.roundRect(36, 742, 110, 36, 4, stroke=0, fill=1)
    
    c.setFont('ArialUnicode', 11)
    c.setFillColor(colors.white)
    c.drawString(46, 762, "JLPT N3")
    c.setFont('ArialUnicode', 8.5)
    c.drawString(46, 749, "Kanji Mind Map")
    
    c.setFillColor(HEADER_TEXT_MAIN)
    c.setFont('ArialUnicode', 13)
    c.drawString(158, 762, "JLPT N3 Kanji Mind Map — Weeks 1–6")
    
    c.setFillColor(HEADER_TEXT_SUB)
    c.setFont('ArialUnicode', 9)
    c.drawString(158, 748, "336 Kanji arranged into 104 Component Radical Trees")
    
    c.setStrokeColor(colors.HexColor('#CBD5E1'))
    c.setLineWidth(0.8)
    c.line(36, 734, 576, 734)
    c.line(36, 36, 576, 36)
    
    c.setFont('ArialUnicode', 8)
    c.setFillColor(colors.HexColor('#64748B'))
    c.drawString(36, 22, "kanji60s.com • JLPT N3 Mind Map Flowcharts")
    c.drawRightString(576, 22, f"Page {page_num} of {total_pages}")

def draw_kanji_node(c, x, y, item):
    kanji = item['kanji']
    reading = item['reading']
    meaning = item['meaning']
    
    w_box, h_box = 28, 28
    x_box = x - w_box / 2.0
    y_box = y - h_box / 2.0
    
    # Meaning lines above box
    m_lines = wrap_text_multiline(meaning, 16)
    c.setFont('ArialUnicode', 7)
    c.setFillColor(MEANING_COLOR)
    # Draw bottom-up from box top
    for i, line in enumerate(reversed(m_lines)):
        c.drawCentredString(x, y_box + h_box + 3 + i * 8.5, line)
        
    # Kanji Box
    c.setFillColor(KANJI_BOX_BG)
    c.setStrokeColor(KANJI_BOX_BORDER)
    c.setLineWidth(0.75)
    c.roundRect(x_box, y_box, w_box, h_box, 3, stroke=1, fill=1)
    
    # Kanji Char
    c.setFont('ArialUnicode', 14)
    c.setFillColor(KANJI_CHAR_COLOR)
    c.drawCentredString(x, y_box + 7, kanji)
    
    # Reading lines below box
    r_lines = format_reading(reading, 16)
    c.setFont('ArialUnicode', 6.2)
    c.setFillColor(READING_COLOR)
    for i, line in enumerate(r_lines):
        c.drawCentredString(x, y_box - 9 - i * 7.5, line)

def calculate_group_layout(group):
    items = group['items']
    N = len(items)
    if N <= 6:
        num_cols = 1
    elif N <= 14:
        num_cols = 2
    else:
        num_cols = 3
    items_per_col = math.ceil(N / num_cols)
    return num_cols, items_per_col

def generate_pdf(output_filename="kflow_fixed.pdf"):
    groups = parse_groups(RAW_DATA)
    c = canvas.Canvas(output_filename, pagesize=letter)
    
    top_y = 715
    bottom_y = 50
    page_height_available = top_y - bottom_y
    
    row_height = 76
    group_padding = 24
    
    pages_data = []
    current_page_groups = []
    current_page_height = 0
    
    for g in groups:
        num_cols, rows_per_col = calculate_group_layout(g)
        g_height = max(rows_per_col * row_height + 10, 58)
        needed_height = g_height + (group_padding if current_page_groups else 0)
        
        if current_page_groups and (current_page_height + needed_height > page_height_available):
            pages_data.append(current_page_groups)
            current_page_groups = [(g, g_height, num_cols, rows_per_col)]
            current_page_height = g_height
        else:
            if current_page_groups:
                current_page_height += group_padding
            current_page_groups.append((g, g_height, num_cols, rows_per_col))
            current_page_height += g_height
            
    if current_page_groups:
        pages_data.append(current_page_groups)
        
    total_pages = len(pages_data)
    
    for page_idx, page_groups in enumerate(pages_data):
        page_num = page_idx + 1
        draw_header_footer(c, page_num, total_pages)
        curr_y = top_y - 10
        
        for g, g_height, num_cols, rows_per_col in page_groups:
            items = g['items']
            N = len(items)
            
            group_top_y = curr_y
            group_center_y = group_top_y - (g_height / 2.0)
            
            x_root = 42
            w_root = 100
            h_root = 32
            y_root = group_center_y - (h_root / 2.0)
            
            # Root box
            c.setFillColor(RADICAL_BOX_BG)
            c.setStrokeColor(colors.HexColor('#0F172A'))
            c.setLineWidth(0.8)
            c.roundRect(x_root, y_root, w_root, h_root, 4, stroke=1, fill=1)
            
            c.setFillColor(RADICAL_BOX_TXT)
            c.setFont('ArialUnicode', 9)
            c.drawCentredString(x_root + w_root/2.0, y_root + 18, g['title'])
            c.setFont('ArialUnicode', 7.5)
            c.setFillColor(colors.HexColor('#94A3B8'))
            c.drawCentredString(x_root + w_root/2.0, y_root + 6, f"[{N} kanji]")
            
            root_connect_x = x_root + w_root
            root_connect_y = group_center_y
            
            if num_cols == 1:
                col_x_offsets = [260]
            elif num_cols == 2:
                col_x_offsets = [220, 410]
            else:
                col_x_offsets = [195, 340, 485]
                
            # Build column matrix of positions and items
            columns_data = []
            for col in range(num_cols):
                col_items = items[col * rows_per_col : (col + 1) * rows_per_col]
                start_y_col = group_top_y - 30
                nodes_info = []
                for r_idx, item in enumerate(col_items):
                    y_node = start_y_col - (r_idx * row_height)
                    nodes_info.append((col_x_offsets[col], y_node, item))
                columns_data.append(nodes_info)
                
            # DRAW LINES (BEFORE NODES)
            c.setStrokeColor(LINE_COLOR)
            c.setLineWidth(0.9)
            
            # 1. Root to Column 0 nodes
            for x_n, y_n, _ in columns_data[0]:
                c.line(root_connect_x, root_connect_y, x_n - 14, y_n)
                
            # 2. Column 0 to Column 1 nodes (paired)
            if num_cols >= 2:
                for i in range(min(len(columns_data[0]), len(columns_data[1]))):
                    x0, y0, _ = columns_data[0][i]
                    x1, y1, _ = columns_data[1][i]
                    c.line(x0 + 14, y0, x1 - 14, y1)
                    
            # 3. Column 1 to Column 2 nodes (paired)
            if num_cols == 3:
                for i in range(min(len(columns_data[1]), len(columns_data[2]))):
                    x1, y1, _ = columns_data[1][i]
                    x2, y2, _ = columns_data[2][i]
                    c.line(x1 + 14, y1, x2 - 14, y2)
                    
            # DRAW KANJI NODES
            for col_info in columns_data:
                for x_n, y_n, item in col_info:
                    draw_kanji_node(c, x_n, y_n, item)
                    
            curr_y -= (g_height + group_padding)
            
        c.showPage()
    c.save()
    print(f"Generated {output_filename} successfully!")

if __name__ == '__main__':
    generate_pdf("kflow_fixed.pdf")

import os
import re
import math
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

FONT_PATH = '/System/Library/Fonts/Supplemental/Arial Unicode.ttf'
pdfmetrics.registerFont(TTFont('ArialUnicode', FONT_PATH))

# Let's test a clean, non-crossing tree renderer with full readings and full meanings!

c = canvas.Canvas("test_full_v3.pdf", pagesize=letter)

# Header
c.setFillColor(colors.HexColor('#1E8449'))
c.roundRect(36, 742, 110, 36, 4, stroke=0, fill=1)
c.setFont('ArialUnicode', 11)
c.setFillColor(colors.white)
c.drawString(46, 762, "JLPT N3")
c.setFont('ArialUnicode', 8.5)
c.drawString(46, 749, "Kanji Mind Map")

c.setFillColor(colors.HexColor('#0F172A'))
c.setFont('ArialUnicode', 13)
c.drawString(158, 762, "JLPT N3 Kanji Mind Map — Weeks 1–6")
c.setFillColor(colors.HexColor('#475569'))
c.setFont('ArialUnicode', 9)
c.drawString(158, 748, "Full Readings & Meanings • Zero-Crossing Mind Map Trees")

c.setStrokeColor(colors.HexColor('#CBD5E1'))
c.setLineWidth(0.8)
c.line(36, 734, 576, 734)

# Sample items from 糸 (thread 糸) [13 kanji]
ito_items = [
    {"kanji": "線", "reading": "sen", "meaning": "line"},
    {"kanji": "約", "reading": "yaku", "meaning": "promise, approximately"},
    {"kanji": "絵", "reading": "kai / e", "meaning": "picture, painting"},
    {"kanji": "組", "reading": "kumi, ku(mu)", "meaning": "group, assemble"},
    {"kanji": "続", "reading": "zoku / tsuzu(ku), tsuzu(keru)", "meaning": "continue"},
    {"kanji": "級", "reading": "kyū", "meaning": "grade, class"},
    {"kanji": "紅", "reading": "kō / beni", "meaning": "crimson, red"},
    {"kanji": "緑", "reading": "ryoku / midori", "meaning": "green"},
    {"kanji": "結", "reading": "ketsu / musu(bu)", "meaning": "tie, connect"},
    {"kanji": "経", "reading": "kei", "meaning": "pass through, manage"},
    {"kanji": "練", "reading": "ren", "meaning": "practice, train"},
    {"kanji": "細", "reading": "hoso(i), koma(kai)", "meaning": "thin, detailed"},
    {"kanji": "絡", "reading": "raku", "meaning": "entangle, contact"}
]

# Function to draw multiline text cleanly centered
def draw_multiline_center(c, x, y_start, text, font_name, font_size, color, leading=9, max_width=110):
    c.setFont(font_name, font_size)
    c.setFillColor(color)
    
    # Split text into lines if longer than max_width or has comma/slash splits
    words = text.split(' ')
    lines = []
    curr_line = ""
    for w in words:
        test_line = (curr_line + " " + w).strip()
        if c.stringWidth(test_line, font_name, font_size) <= max_width:
            curr_line = test_line
        else:
            if curr_line:
                lines.append(curr_line)
            curr_line = w
    if curr_line:
        lines.append(curr_line)
        
    for i, line in enumerate(lines):
        c.drawCentredString(x, y_start - (i * leading), line)
    return len(lines)

def draw_kanji_node_full(c, x, y, item):
    kanji = item['kanji']
    reading = f"[{item['reading']}]"
    meaning = item['meaning']
    
    w_box, h_box = 28, 28
    x_box = x - w_box / 2.0
    y_box = y - h_box / 2.0
    
    # 1. Meaning text above box (multiline if needed)
    # Check lines
    words = meaning.split(' ')
    m_lines = []
    cur = ""
    for w in words:
        if c.stringWidth((cur + " " + w).strip(), 'ArialUnicode', 7.5) <= 100:
            cur = (cur + " " + w).strip()
        else:
            if cur: m_lines.append(cur)
            cur = w
    if cur: m_lines.append(cur)
    
    # Draw meaning lines going upwards from y_box + h_box + 3
    c.setFont('ArialUnicode', 7.5)
    c.setFillColor(colors.HexColor('#1E293B'))
    for idx_l, line in enumerate(reversed(m_lines)):
        c.drawCentredString(x, y_box + h_box + 3 + (idx_l * 9), line)
        
    # 2. Kanji Box
    c.setFillColor(colors.HexColor('#F8FAFC'))
    c.setStrokeColor(colors.HexColor('#475569'))
    c.setLineWidth(0.75)
    c.roundRect(x_box, y_box, w_box, h_box, 3, stroke=1, fill=1)
    
    c.setFont('ArialUnicode', 14)
    c.setFillColor(colors.HexColor('#8C2A1E'))
    c.drawCentredString(x, y_box + 7, kanji)
    
    # 3. Reading text below box (multiline if needed)
    r_words = reading.split(' ')
    r_lines = []
    cur = ""
    for w in r_words:
        if c.stringWidth((cur + " " + w).strip(), 'ArialUnicode', 6.5) <= 105:
            cur = (cur + " " + w).strip()
        else:
            if cur: r_lines.append(cur)
            cur = w
    if cur: r_lines.append(cur)
    
    c.setFont('ArialUnicode', 6.5)
    c.setFillColor(colors.HexColor('#4B5563'))
    for idx_l, line in enumerate(r_lines):
        c.drawCentredString(x, y_box - 10 - (idx_l * 8), line)

# Render 糸 group with ZERO line crossing!
# Col 1: items 0..6 (7 items)
# Col 2: items 7..12 (6 items)
# Tree structure: Root connects to Col 1 items. Each Col 1 item connects horizontally to corresponding Col 2 item!

x_root = 42
w_root = 90
h_root = 32
group_top = 680
row_pitch = 65

col1_items = ito_items[0:7]
col2_items = ito_items[7:13]

num_rows = len(col1_items)
y_center = group_top - ((num_rows - 1) * row_pitch) / 2.0
y_root_box = y_center - h_root / 2.0

# Draw Root
c.setFillColor(colors.HexColor('#1E293B'))
c.roundRect(x_root, y_root_box, w_root, h_root, 4, stroke=1, fill=1)
c.setFillColor(colors.white)
c.setFont('ArialUnicode', 9.5)
c.drawCentredString(x_root + w_root/2.0, y_root_box + 18, "糸 (thread 糸)")
c.setFont('ArialUnicode', 7.5)
c.setFillColor(colors.HexColor('#94A3B8'))
c.drawCentredString(x_root + w_root/2.0, y_root_box + 6, "[13 kanji]")

x_col1 = 210
x_col2 = 410

# 1. Connect Root to Col 1
for i, item in enumerate(col1_items):
    y_n = group_top - (i * row_pitch)
    
    # Line from Root right edge to Col 1 node left edge
    c.setStrokeColor(colors.HexColor('#94A3B8'))
    c.setLineWidth(0.9)
    c.line(x_root + w_root, y_center, x_col1 - 14, y_n)
    
    draw_kanji_node_full(c, x_col1, y_n, item)

# 2. Connect Col 1 items to Col 2 items directly horizontally!
for j, item in enumerate(col2_items):
    y_n2 = group_top - (j * row_pitch)
    
    # Line from Col 1 right edge to Col 2 left edge
    c.setStrokeColor(colors.HexColor('#94A3B8'))
    c.setLineWidth(0.9)
    c.line(x_col1 + 14, y_n2, x_col2 - 14, y_n2)
    
    draw_kanji_node_full(c, x_col2, y_n2, item)

c.save()
print("Saved test_full_v3.pdf")

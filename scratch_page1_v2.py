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

# Let's test page 1 layout with 水 (25 kanji) split into 3 clean side-by-side sub-trees!

c = canvas.Canvas("test_page1_v2.pdf", pagesize=letter)

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
c.drawString(158, 748, "Group 1: 水 (water 氵) [25 kanji]")

c.setStrokeColor(colors.HexColor('#CBD5E1'))
c.setLineWidth(0.8)
c.line(36, 734, 576, 734)

# Items for 水
water_items = [
    {"kanji": "満", "reading": "man", "meaning": "full, satisfied"},
    {"kanji": "港", "reading": "kō / minato", "meaning": "harbor, port"},
    {"kanji": "準", "reading": "jun", "meaning": "prepare, standard"},
    {"kanji": "温", "reading": "on / atata(kai)", "meaning": "warm, temp."},
    {"kanji": "法", "reading": "hō", "meaning": "law, method"},
    {"kanji": "減", "reading": "gen / he(ru)...", "meaning": "decrease"},
    {"kanji": "湯", "reading": "yu", "meaning": "hot water"},
    {"kanji": "氷", "reading": "kōri", "meaning": "ice"},
    {"kanji": "混", "reading": "kon / ma(zeru)", "meaning": "mix"},
    {"kanji": "渡", "reading": "wata(ru)...", "meaning": "cross, hand over"},
    {"kanji": "濃", "reading": "ko(i)", "meaning": "thick, dark"},
    {"kanji": "消", "reading": "shō / ki(eru)...", "meaning": "extinguish"},
    {"kanji": "汗", "reading": "ase", "meaning": "sweat"},
    {"kanji": "涙", "reading": "namida", "meaning": "tear"},
    {"kanji": "治", "reading": "ji/chi / nao...", "meaning": "cure, heal"},
    {"kanji": "汚", "reading": "kitanai...", "meaning": "dirty"},
    {"kanji": "泣", "reading": "na(ku)", "meaning": "cry"},
    {"kanji": "泊", "reading": "haku / to...", "meaning": "stay overnight"},
    {"kanji": "波", "reading": "ha / nami", "meaning": "wave"},
    {"kanji": "決", "reading": "ketsu / ki...", "meaning": "decide"},
    {"kanji": "済", "reading": "sai / su(mu)", "meaning": "finish, settle"},
    {"kanji": "活", "reading": "katsu", "meaning": "life, activity"},
    {"kanji": "流", "reading": "naga(reru)...", "meaning": "flow"},
    {"kanji": "油", "reading": "yu / abura", "meaning": "oil"},
    {"kanji": "湖", "reading": "ko / mizuumi", "meaning": "lake"}
]

# Split into 3 columns: 9, 8, 8 items
cols_data = [
    ("水 • Set A", water_items[0:9]),
    ("水 • Set B", water_items[9:17]),
    ("水 • Set C", water_items[17:25])
]

# X positions for 3 sub-trees
# Available width = 540 (36..576)
# Each column gets ~170 pt width
col_x_roots = [40, 220, 400]
col_x_nodes = [145, 325, 505]

row_pitch = 60
start_y = 670

for col_idx, (col_title, items) in enumerate(cols_data):
    x_root = col_x_roots[col_idx]
    x_node = col_x_nodes[col_idx]
    
    num_items = len(items)
    tree_height = (num_items - 1) * row_pitch
    y_center = start_y - (tree_height / 2.0)
    
    # Draw Sub-Root Node
    w_root, h_root = 75, 28
    y_root_box = y_center - h_root / 2.0
    c.setFillColor(colors.HexColor('#1E293B'))
    c.roundRect(x_root, y_root_box, w_root, h_root, 4, stroke=1, fill=1)
    
    c.setFillColor(colors.white)
    c.setFont('ArialUnicode', 8.5)
    c.drawCentredString(x_root + w_root/2.0, y_root_box + 17, col_title)
    c.setFillColor(colors.HexColor('#94A3B8'))
    c.setFont('ArialUnicode', 7.0)
    c.drawCentredString(x_root + w_root/2.0, y_root_box + 6, f"[{num_items} kanji]")
    
    root_px = x_root + w_root
    root_py = y_center
    
    # Draw Kanji Nodes
    for i, item in enumerate(items):
        y_n = start_y - (i * row_pitch)
        
        # Line from root to child
        c.setStrokeColor(colors.HexColor('#94A3B8'))
        c.setLineWidth(0.9)
        c.line(root_px, root_py, x_node - 14, y_n)
        
        # Draw node
        kanji = item['kanji']
        reading = f"[{item['reading']}]"
        meaning = item['meaning']
        
        # Meaning above box
        c.setFont('ArialUnicode', 7.5)
        c.setFillColor(colors.HexColor('#1E293B'))
        c.drawCentredString(x_node, y_n + 17, meaning)
        
        # Box
        w_b, h_b = 28, 28
        c.setFillColor(colors.HexColor('#F8FAFC'))
        c.setStrokeColor(colors.HexColor('#475569'))
        c.setLineWidth(0.75)
        c.roundRect(x_node - w_b/2.0, y_n - h_b/2.0, w_b, h_b, 3, stroke=1, fill=1)
        
        # Kanji character inside
        c.setFont('ArialUnicode', 14)
        c.setFillColor(colors.HexColor('#8C2A1E'))
        c.drawCentredString(x_node, y_n - 4, kanji)
        
        # Reading below box
        c.setFont('ArialUnicode', 6.5)
        c.setFillColor(colors.HexColor('#4B5563'))
        c.drawCentredString(x_node, y_n - 22, reading)

c.save()
print("Generated test_page1_v2.pdf")

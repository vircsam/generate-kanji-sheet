import os
import re
import math
import subprocess
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# 1. Register Font
FONT_PATH = '/System/Library/Fonts/Supplemental/Arial Unicode.ttf'
pdfmetrics.registerFont(TTFont('ArialUnicode', FONT_PATH))

# 2. RAW DATA
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

# Palette
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
    
    m_lines = wrap_text_multiline(meaning, 16)
    c.setFont('ArialUnicode', 7)
    c.setFillColor(MEANING_COLOR)
    for i, line in enumerate(reversed(m_lines)):
        c.drawCentredString(x, y_box + h_box + 3 + i * 8.5, line)
        
    c.setFillColor(KANJI_BOX_BG)
    c.setStrokeColor(KANJI_BOX_BORDER)
    c.setLineWidth(0.75)
    c.roundRect(x_box, y_box, w_box, h_box, 3, stroke=1, fill=1)
    
    c.setFont('ArialUnicode', 14)
    c.setFillColor(KANJI_CHAR_COLOR)
    c.drawCentredString(x, y_box + 7, kanji)
    
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

def generate_pdf(output_filename="kflow.pdf"):
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
            
            # Root to Col 0
            for x_n, y_n, _ in columns_data[0]:
                c.line(root_connect_x, root_connect_y, x_n - 14, y_n)
                
            # Col 0 to Col 1
            if num_cols >= 2:
                for i in range(min(len(columns_data[0]), len(columns_data[1]))):
                    x0, y0, _ = columns_data[0][i]
                    x1, y1, _ = columns_data[1][i]
                    c.line(x0 + 14, y0, x1 - 14, y1)
                    
            # Col 1 to Col 2
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
    print(f"Generated {output_filename} successfully ({total_pages} pages).")
    return total_pages

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def set_table_borders(table, color="CBD5E1"):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
            f'<w:left w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
            f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
            f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def generate_docx(pdf_filename="kflow.pdf", docx_filename="kflow.docx"):
    # 1. Render PDF pages to PNG using sips
    temp_dir = "kflow_pages_temp"
    os.makedirs(temp_dir, exist_ok=True)
    
    # Run sips to split PDF into PNGs
    cmd = f"sips -s format png '{pdf_filename}' --out '{temp_dir}/page.png'"
    subprocess.run(cmd, shell=True, check=True)
    
    # List generated png files sorted
    png_files = sorted([os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.endswith('.png')])
    if not png_files:
        # Fallback if sips output single file or numbered
        png_files = [f for f in os.listdir(temp_dir) if f.endswith('.png')]
        
    doc = Document()
    
    # Configure 0.4 in margins
    for s in doc.sections:
        s.top_margin = Inches(0.4)
        s.bottom_margin = Inches(0.4)
        s.left_margin = Inches(0.4)
        s.right_margin = Inches(0.4)

    # Document Header Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(12)
    p_title.paragraph_format.space_after = Pt(4)
    run_t = p_title.add_run("JLPT N3 Kanji Mind Map Flowcharts")
    run_t.font.name = "Arial"
    run_t.font.size = Pt(22)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(0x1E, 0x84, 0x49) # Header green

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(18)
    run_sub = p_sub.add_run("336 Kanji arranged into 104 Component Radical Trees (Complete Readings & Meanings)")
    run_sub.font.name = "Arial"
    run_sub.font.size = Pt(11)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    # Insert Mind Map Flowchart images per page
    for idx, png_path in enumerate(png_files):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(4)
        p_img.paragraph_format.space_after = Pt(8)
        run_img = p_img.add_run()
        run_img.add_picture(png_path, width=Inches(7.5))
        
        if idx < len(png_files) - 1:
            doc.add_page_break()

    # Save document
    doc.save(docx_filename)
    print(f"Successfully generated Word file: {docx_filename}")

if __name__ == "__main__":
    generate_pdf("kflow.pdf")
    generate_docx("kflow.pdf", "kflow.docx")

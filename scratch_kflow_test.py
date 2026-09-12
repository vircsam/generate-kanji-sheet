import os
import re
import math
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

# Register Font
FONT_PATH = '/System/Library/Fonts/Supplemental/Arial Unicode.ttf'
pdfmetrics.registerFont(TTFont('ArialUnicode', FONT_PATH))

from kflow import RAW_DATA, parse_groups

def wrap_text(text, max_chars=16):
    text = text.strip()
    if len(text) <= max_chars:
        return [text]
    
    # Check splitters
    for sep in [' / ', '/', ', ', ',']:
        if sep in text:
            parts = text.split(sep)
            mid = len(parts) // 2
            sep_join = sep.strip() + " " if sep.strip() in [',', '/'] else sep
            line1 = sep_join.join(parts[:mid]).strip()
            line2 = sep_join.join(parts[mid:]).strip()
            if len(line1) <= max_chars * 1.5 and len(line2) <= max_chars * 1.5:
                return [line1, line2]
                
    # Fallback word split
    words = text.split()
    line1, line2 = "", ""
    for w in words:
        if len((line1 + " " + w).strip()) <= max_chars or not line1:
            line1 = (line1 + " " + w).strip()
        else:
            line2 = (line2 + " " + w).strip()
    if line2:
        return [line1, line2]
    return [text]

print("Text wrapping test:")
print("extinguish, disappear ->", wrap_text("extinguish, disappear", 14))
print("ji/chi / nao(ru), nao(su) ->", wrap_text("ji/chi / nao(ru), nao(su)", 14))
print("hoso(i), koma(kai) ->", wrap_text("hoso(i), koma(kai)", 14))

import os
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.pdfgen import canvas

FONT_PATH = '/System/Library/Fonts/Supplemental/Arial Unicode.ttf'

try:
    pdfmetrics.registerFont(TTFont('ArialUnicode', FONT_PATH))
    print("ArialUnicode registered successfully!")
except Exception as e:
    print("Error registering font:", e)

# Test creating a small PDF
c = canvas.Canvas("test_font.pdf")
c.setFont('ArialUnicode', 14)
c.drawString(100, 700, "水 (water 氵) [25 kanji]")
c.drawString(100, 670, "満 [man] full, satisfied")
c.drawString(100, 640, "港 [kō / minato] harbor, port")
c.save()
print("Saved test_font.pdf")

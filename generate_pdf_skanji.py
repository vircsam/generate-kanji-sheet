import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from soma import (
    CHAPTER_1_DATA,
    CHAPTER_2_DATA,
    CHAPTER_3_DATA,
    CHAPTER_4_DATA,
    CHAPTER_5_DATA,
    CHAPTER_6_DATA
)

# Register Arial Unicode font for full CJK character support
FONT_PATH = '/System/Library/Fonts/Supplemental/Arial Unicode.ttf'
pdfmetrics.registerFont(TTFont('ArialUnicode', FONT_PATH))

def create_skanji_pdf(datasets_dict, output_pdf_path):
    # Setup document with 0.4 in (28.8 pt) margins
    doc = SimpleDocTemplate(
        output_pdf_path,
        pagesize=letter,
        leftMargin=28.8,
        rightMargin=28.8,
        topMargin=28.8,
        bottomMargin=28.8
    )

    styles = getSampleStyleSheet()

    # Define custom styles using registered ArialUnicode font
    chapter_title_style = ParagraphStyle(
        'ChapterTitle',
        parent=styles['Normal'],
        fontName='ArialUnicode',
        fontSize=16,
        leading=20,
        textColor=colors.HexColor('#8C2A1E'),
        spaceBefore=6,
        spaceAfter=8,
        keepWithNext=True
    )

    hdr_left_style = ParagraphStyle(
        'HdrLeft',
        parent=styles['Normal'],
        fontName='ArialUnicode',
        fontSize=9,
        leading=11,
        textColor=colors.HexColor('#555555'),
        alignment=0 # Left
    )

    hdr_center_style = ParagraphStyle(
        'HdrCenter',
        parent=styles['Normal'],
        fontName='ArialUnicode',
        fontSize=9,
        leading=11,
        textColor=colors.HexColor('#555555'),
        alignment=1 # Center
    )

    kanji_style = ParagraphStyle(
        'KanjiText',
        parent=styles['Normal'],
        fontName='ArialUnicode',
        fontSize=20,
        leading=23,
        textColor=colors.HexColor('#111111')
    )

    meaning_style = ParagraphStyle(
        'MeaningText',
        parent=styles['Normal'],
        fontName='ArialUnicode',
        fontSize=8.5,
        leading=10.5,
        textColor=colors.HexColor('#555555')
    )

    story = []

    # Printable width = 8.5 * 72 - 2 * 28.8 = 554.4 pt
    col_widths = [194.4, 180.0, 180.0]

    for set_idx, (set_name, kanji_list) in enumerate(datasets_dict.items()):
        # Chapter Title
        story.append(Paragraph(f"<b>{set_name}</b>", chapter_title_style))
        story.append(Spacer(1, 4))

        table_data = []

        # Header Row
        hdr_row = [
            Paragraph("<b>KANJI & MEANING</b>", hdr_left_style),
            Paragraph("<b>PRACTICE 1</b>", hdr_center_style),
            Paragraph("<b>PRACTICE 2</b>", hdr_center_style)
        ]
        table_data.append(hdr_row)

        # Row heights array
        row_heights = [24.0] # Header row height

        for kanji, meaning, _vocab in kanji_list:
            cell_0_content = [
                Paragraph(f"<b>{kanji}</b>", kanji_style),
                Spacer(1, 2),
                Paragraph(f"<i>{meaning}</i>", meaning_style)
            ]
            
            row = [cell_0_content, "", ""]
            table_data.append(row)
            row_heights.append(48.0) # Height for hand-writing practice

        # Create Table
        t = Table(table_data, colWidths=col_widths, rowHeights=row_heights, repeatRows=1)

        t_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#EDE5D8')),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D3C7B6')),
            ('BACKGROUND', (0, 1), (0, -1), colors.HexColor('#F9F6F0')),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ])

        t.setStyle(t_style)
        story.append(t)

        if set_idx < len(datasets_dict) - 1:
            story.append(PageBreak())

    doc.build(story)
    print(f"Successfully generated PDF: {output_pdf_path}")

if __name__ == "__main__":
    combined_datasets = {}
    for d in [CHAPTER_1_DATA, CHAPTER_2_DATA, CHAPTER_3_DATA, CHAPTER_4_DATA, CHAPTER_5_DATA, CHAPTER_6_DATA]:
        combined_datasets.update(d)

    create_skanji_pdf(combined_datasets, "Kanji_Practice_All_Chapters_3cols.pdf")

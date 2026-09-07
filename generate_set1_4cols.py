import json
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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

def set_table_borders(table, color="D3C7B6"):
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

# Chapter 1 (Week 1, Days 1 to 7): 58 Kanji in total
CHAPTER_1_DATA = {
    "Chapter 1 (Week 1)": [
        # 第1週 1日目：駐車場
        ["駐", "チュウ / park", [
            ["駐車", "ちゅうしゃ", "park"],
            ["駐車場", "ちゅうしゃじょう", "a parking lot"]
        ]],
        ["無", "ム / な-い / without, non-existent", [
            ["無休", "むきゅう", "work without a holiday"],
            ["無理(な)", "むり(な)", "unreasonable"],
            ["無料", "むりょう", "free of charge"],
            ["無い", "ない", "does not exist, do not have, gone"]
        ]],
        ["満", "マン / full", [
            ["満車", "まんしゃ", "full (of cars)"],
            ["不満(な)", "ふまん(な)", "dissatisfied"],
            ["満員", "まんいん", "full (of people)"]
        ]],
        ["向", "コウ / む-こう / む-かう / む-き / direction, face", [
            ["方向", "ほうこう", "direction"],
            ["向かう", "むかう", "go forward"],
            ["向こう", "むこう", "over there/beyond"],
            ["〇〇向き", "〇〇むき", "be suitable for 〇〇"]
        ]],
        ["禁", "キン / prohibition", [
            ["禁止", "きんし", "prohibition"]
        ]],
        ["関", "カン / connection, relation", [
            ["関心", "かんしん", "interest"],
            ["関する", "かんする", "related"]
        ]],
        ["係", "ケイ / かかり / duty, person in charge", [
            ["関係", "かんけい", "relation/connection"],
            ["係", "かかり", "person in charge"]
        ]],
        ["断", "ダン / ことわ-る / refuse, cut off", [
            ["無断", "むだん", "without permission"],
            ["断る", "ことわる", "refuse"],
            ["断水", "だんすい", "suspension of the water supply"]
        ]],

        # 第1週 2日目：横断歩道
        ["横", "オウ / よこ / side, horizontal", [
            ["横断", "おうだん", "a crossing"],
            ["横", "よこ", "side"],
            ["横断歩道", "おうだんほどう", "a pedestrian crossing"]
        ]],
        ["押", "お-す / お-さえる / push, hold down", [
            ["押す", "おす", "push"],
            ["押し入れ", "おしいれ", "a closet"],
            ["押さえる", "おさえる", "hold down"]
        ]],
        ["式", "シキ / formula, ceremony, style", [
            ["押しボタン式", "おしボタンしき", "a push-button ..."],
            ["数式", "すうしき", "a numerical formula"],
            ["入学式", "にゅうがくしき", "a ceremony to begin the school term"]
        ]],
        ["信", "シン / trust, signal, belief", [
            ["送信", "そうしん", "transmit"],
            ["自信", "じしん", "confidence"],
            ["信じる", "しんじる", "believe"],
            ["信用", "しんよう", "trust"]
        ]],
        ["号", "ゴウ / number, signal", [
            ["信号", "しんごう", "a signal / a traffic light"],
            ["～号車", "～ごうしゃ", "carriage number ..."]
        ]],
        ["確", "カク / たし-か / たし-かめる / certain, confirm", [
            ["正確(な)", "せいかく(な)", "correct, accurate"],
            ["確かめる", "たしかめる", "confirm/verify"],
            ["確か(な)", "たしか(な)", "certain"]
        ]],
        ["認", "ニン / みと-める / recognize, approve", [
            ["確認", "かくにん", "confirmation"],
            ["認める", "みとめる", "admit/approve"]
        ]],
        ["飛", "ヒ / と-ぶ / fly", [
            ["飛行場", "ひこうじょう", "an airport"],
            ["飛ぶ", "とぶ", "fly"]
        ]],

        # 第1週 3日目：サイン
        ["非", "ヒ / non-, emergency", [
            ["非常(の)", "ひじょう(の)", "emergency"],
            ["非常口", "ひじょうぐち", "an emergency exit"],
            ["非常に", "ひじょうに", "extremely"]
        ]],
        ["常", "ジョウ / normal, usual", [
            ["日常(の)", "にちじょう(の)", "usual, everyday"],
            ["正常(な)", "せいじょう(な)", "normal"]
        ]],
        ["階", "カイ / floor", [
            ["～階", "～かい", "... floor"]
        ]],
        ["段", "ダン / step, grade", [
            ["階段", "かいだん", "stairs"]
        ]],
        ["箱", "はこ / box", [
            ["箱", "はこ", "a box"],
            ["ごみ箱", "ごみばこ", "a trash box"]
        ]],
        ["危", "キ / あぶ-ない / dangerous", [
            ["危険(な)", "きけん(な)", "danger"],
            ["危ない", "あぶない", "dangerous"]
        ]],
        ["険", "ケン / steep, risk, danger", [
            ["危険(な)", "きけん(な)", "danger"]
        ]],
        ["捨", "す-てる / throw away", [
            ["捨てる", "すてる", "throw away"]
        ]],

        # 第1週 4日目：駅のホーム
        ["線", "セン / line, track", [
            ["線", "せん", "a line"],
            ["～番線", "～ばんせん", "line (platform) number ..."]
        ]],
        ["面", "メン / surface, face, screen", [
            ["全面", "ぜんめん", "whole"],
            ["〇〇方面", "〇〇ほうめん", "... area"],
            ["画面", "がめん", "a screen"]
        ]],
        ["普", "フ / general, ordinary", [
            ["普通(の)", "ふつう(の)", "ordinary"]
        ]],
        ["各", "カク / each, every", [
            ["各駅", "かくえき", "every station"],
            ["各自", "かくじ", "respective / each person"],
            ["各国", "かっこく", "each country"]
        ]],
        ["次", "ジ / つぎ / next, order", [
            ["目次", "もくじ", "contents"],
            ["次", "つぎ", "next"],
            ["次回", "じかい", "the next time"]
        ]],
        ["快", "カイ / pleasant, fast", [
            ["快速", "かいそく", "ordinary"]
        ]],
        ["速", "ソク / はや-い / fast, speed", [
            ["高速道路", "こうそくどうろ", "a highway"],
            ["速い", "はやい", "fast"],
            ["速度", "そくど", "speed"]
        ]],
        ["過", "カ / す-ぎる / pass, exceed, past", [
            ["通過", "つうか", "passage / transit"],
            ["過ぎる", "すぎる", "pass"],
            ["過去", "かこ", "past"]
        ]],
        ["鉄", "テツ / iron, steel, railway", [
            ["地下鉄", "ちかてつ", "a subway"],
            ["鉄", "てつ", "steel"],
            ["鉄道", "てつどう", "a railway"]
        ]],

        # 第1週 5日目：特急電車
        ["指", "シ / ゆび / finger, designate", [
            ["指定", "してい", "specify, designate"],
            ["指", "ゆび", "a finger"],
            ["指定席", "していせき", "a reserved seat"],
            ["指輪", "ゆびわ", "a ring"]
        ]],
        ["定", "テイ / decide, fix", [
            ["定休日", "ていきゅうび", "a set/regular holiday"],
            ["安定", "あんてい", "stable"],
            ["不安定", "ふあんてい", "unstable"]
        ]],
        ["席", "セキ / seat", [
            ["席", "せき", "a seat"],
            ["出席", "しゅっせき", "attend"],
            ["欠席", "けっせき", "absence"]
        ]],
        ["由", "ユウ / reason, origin", [
            ["自由(な)", "じゆう(な)", "free"],
            ["理由", "りゆう", "reason"],
            ["自由席", "じゆうせき", "a non-reserved seat"]
        ]],
        ["番", "バン / number, turn", [
            ["番号", "ばんごう", "a number"],
            ["～番線", "～ばんせん", "line (platform) number ..."],
            ["～番", "～ばん", "number ..."]
        ]],
        ["窓", "まど / window", [
            ["窓", "まど", "a window"],
            ["窓口", "まどぐち", "a teller's window"]
        ]],
        ["側", "がわ / side", [
            ["両側", "りょうがわ", "both sides"],
            ["窓側", "まどがわ", "a window seat"],
            ["右側", "みぎがわ", "the right side"]
        ]],
        ["路", "ロ / road, path", [
            ["通路", "つうろ", "an aisle"],
            ["線路", "せんろ", "a railway"],
            ["道路", "どうろ", "a road"]
        ]],

        # 第1週 6日目：バス
        ["停", "テイ / stop", [
            ["停車", "ていしゃ", "stop a vehicle"],
            ["バス停", "バスてい", "a bus stop"]
        ]],
        ["整", "セイ / organize, arrange", [
            ["整理", "せいり", "tidy"],
            ["整理券", "せいりけん", "numbered ticket"]
        ]],
        ["券", "ケン / ticket", [
            ["駐車券", "ちゅうしゃけん", "a parking ticket"],
            ["回数券", "かいすうけん", "(book of) commuter ticket(s)"],
            ["乗車券", "じょうしゃけん", "a (boarding) ticket"]
        ]],
        ["現", "ゲン / あらわ-れる / appear, cash", [
            ["現金", "げんきん", "cash"],
            ["現れる", "あらわれる", "appear"],
            ["表現", "ひょうげん", "expression"]
        ]],
        ["両", "リョウ / both, car", [
            ["両親", "りょうしん", "parents"],
            ["～両", "～りょう", "... cars on a train"]
        ]],
        ["替", "か-える / exchange, change", [
            ["取り替える", "とりかえる", "exchange"],
            ["着替える", "きがえる", "change your clothes"],
            ["両替", "りょうがえ", "exchange"]
        ]],
        ["優", "ユウ / やさ-しい / priority, kind", [
            ["優先席", "ゆうせんせき", "a priority seat"],
            ["優しい", "やさしい", "kind"],
            ["女優", "じょゆう", "an actress"]
        ]],
        ["座", "ザ / すわ-る / seat, sit", [
            ["座席", "ざせき", "a seat"],
            ["座る", "すわる", "sit"],
            ["正座", "せいざ", "sit on the floor Japanese style"]
        ]],
        ["降", "コウ / お-りる / ふ-る / get off, fall", [
            ["降車口", "こうしゃぐち", "exit (for getting off)"],
            ["降りる", "おりる", "get off"],
            ["以降", "いこう", "after ..."],
            ["降る", "ふる", "fall"]
        ]],

        # 第1週 7日目：クイズ① どちらの字？
        ["未", "ミ / not yet, future", [
            ["未定", "みてい", "undecided/indefinite"],
            ["～未満", "～みまん", "less than ..."],
            ["未来", "みらい", "future"]
        ]],
        ["末", "マツ / end", [
            ["週末", "しゅうまつ", "the weekend"],
            ["年末", "ねんまつ", "the end of the year"],
            ["月末", "げつまつ", "the end of the month"]
        ]],
        ["若", "わか-い / young", [
            ["若い", "わかい", "young"]
        ]],
        ["晩", "バン / night, evening", [
            ["晩", "ばん", "night"],
            ["晩ご飯", "ばんごはん", "dinner"],
            ["今晩", "こんばん", "tonight"],
            ["毎晩", "まいばん", "every night"]
        ]],
        ["島", "トウ / しま / island", [
            ["〇〇島", "〇〇とう", "〇〇 Island"],
            ["島", "しま", "an island"]
        ]],
        ["皿", "さら / plate", [
            ["皿", "さら", "a plate"],
            ["灰皿", "はいざら", "an ashtray"]
        ]],
        ["血", "ケツ / ち / blood", [
            ["出血", "しゅっけつ", "bleed"],
            ["血", "ち", "blood"]
        ]],
        ["助", "ジョ / たす-ける / help, rescue", [
            ["救助", "きゅうじょ", "rescue / help"],
            ["助ける", "たすける", "help"]
        ]]
    ]
}

def generate_kanji_document_4cols(subset_dict, filename="Kanji_Practice_Chapter1_4cols.docx"):
    doc = Document()
    
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(0.4)
        s.bottom_margin = Inches(0.4)
        s.left_margin = Inches(0.4)
        s.right_margin = Inches(0.4)

    # 1 WORD column (2.1 in) + 4 Writing Practice columns (1.4 in each) = 7.7 in total
    COL_WIDTHS = [Inches(2.1)] + [Inches(1.4)] * 4
    
    for set_idx, (set_name, kanji_list) in enumerate(subset_dict.items()):
        p_set = doc.add_paragraph()
        p_set.paragraph_format.space_before = Pt(6)
        p_set.paragraph_format.space_after = Pt(2)
        run_set = p_set.add_run(set_name)
        run_set.font.name = "Georgia"
        run_set.font.size = Pt(16)
        run_set.font.bold = True
        run_set.font.color.rgb = RGBColor(0x8C, 0x2A, 0x1E)

        for kanji, meaning, vocab_list in kanji_list:
            # Header Kanji Banner Table (1x1)
            header_tbl = doc.add_table(rows=1, cols=1)
            header_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            h_cell = header_tbl.rows[0].cells[0]
            h_cell.width = Inches(7.7)
            set_cell_background(h_cell, "F3EFE9")
            set_cell_margins(h_cell, top=40, bottom=40, left=120, right=120)
            
            tcPr = h_cell._element.get_or_add_tcPr()
            borders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:left w:val="single" w:sz="36" w:space="0" w:color="8C2A1E"/>'
                f'<w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(borders)

            hp = h_cell.paragraphs[0]
            hp.paragraph_format.space_before = Pt(1)
            hp.paragraph_format.space_after = Pt(1)
            k_run = hp.add_run(kanji)
            k_run.font.name = "MS Mincho"
            k_run.font.size = Pt(22)
            k_run.font.bold = True

            m_p = h_cell.add_paragraph()
            m_p.paragraph_format.space_before = Pt(0)
            m_p.paragraph_format.space_after = Pt(1)
            m_run = m_p.add_run(meaning)
            m_run.font.name = "Georgia"
            m_run.font.size = Pt(9.5)
            m_run.font.italic = True
            m_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

            doc.add_paragraph().paragraph_format.space_after = Pt(2)

            # Practice Grid Table for this Kanji (5 columns total: WORD + 4 Practice columns)
            table = doc.add_table(rows=1, cols=5)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(table, color="D3C7B6")

            # Header Row
            hdr_row = table.rows[0]
            hdr_titles = ["WORD", "1", "2", "3", "4"]
            for i, cell in enumerate(hdr_row.cells):
                cell.width = COL_WIDTHS[i]
                set_cell_background(cell, "EDE5D8")
                set_cell_margins(cell, top=30, bottom=30, left=40, right=40)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                r = p.add_run(hdr_titles[i])
                r.font.name = "Arial"
                r.font.size = Pt(9)
                r.font.bold = True
                r.font.color.rgb = RGBColor(0x77, 0x77, 0x77)

            # Add 4 practice sub-rows per vocabulary word block
            for w_kanji, w_hira, w_eng in vocab_list:
                r0 = table.add_row()
                r1 = table.add_row()
                r2 = table.add_row()
                r3 = table.add_row()

                for r in [r0, r1, r2, r3]:
                    trPr = r._element.get_or_add_trPr()
                    trHeight = parse_xml(f'<w:trHeight {nsdecls("w")} w:val="450" w:hRule="atLeast"/>')
                    trPr.append(trHeight)
                    for i, c in enumerate(r.cells):
                        c.width = COL_WIDTHS[i]
                        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                # Merge column 0 across r0, r1, r2, r3
                info_cell = r0.cells[0]
                info_cell.merge(r1.cells[0])
                info_cell.merge(r2.cells[0])
                info_cell.merge(r3.cells[0])
                set_cell_margins(info_cell, top=40, bottom=40, left=80, right=40)
                info_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                ip = info_cell.paragraphs[0]
                ip.paragraph_format.space_before = Pt(0)
                ip.paragraph_format.space_after = Pt(1)
                w_run = ip.add_run(w_kanji)
                w_run.font.name = "MS Mincho"
                w_run.font.size = Pt(17)
                w_run.font.bold = True

                hp = info_cell.add_paragraph()
                hp.paragraph_format.space_before = Pt(0)
                hp.paragraph_format.space_after = Pt(1)
                h_run = hp.add_run(w_hira)
                h_run.font.name = "Hiragino Mincho ProN"
                h_run.font.size = Pt(8.5)
                h_run.font.color.rgb = RGBColor(0x8C, 0x6B, 0x3E)

                ep = info_cell.add_paragraph()
                ep.paragraph_format.space_before = Pt(0)
                ep.paragraph_format.space_after = Pt(0)
                e_run = ep.add_run(w_eng)
                e_run.font.name = "Arial"
                e_run.font.size = Pt(7.5)
                e_run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

            # Spacing between Kanji blocks
            sp = doc.add_paragraph()
            sp.paragraph_format.space_before = Pt(0)
            sp.paragraph_format.space_after = Pt(10)

        if set_idx < len(subset_dict) - 1:
            doc.add_page_break()

    doc.save(filename)
    print(f"Successfully generated: {filename}")

if __name__ == "__main__":
    generate_kanji_document_4cols(CHAPTER_1_DATA, "Kanji_Practice_Chapter1_4cols.docx")

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

# Chapter 2 (Week 2, 7 Days): Kanji data
CHAPTER_2_DATA = {
    "Chapter 2 (Week 2)": [
        # 第2週 1日目：レストラン
        ["準", "ジュン / standard, prepare", [
            ["準備", "じゅんび", "preparation"]
        ]],
        ["備", "ビ / そな-える / prepare, provide", [
            ["準備", "じゅんび", "preparation"],
            ["備える", "そなえる", "prepare"]
        ]],
        ["営", "エイ / manage, business", [
            ["営業", "えいぎょう", "business"]
        ]],
        ["閉", "ヘイ / し-まる / し-める / close, shut", [
            ["開閉", "かいへい", "open and shut"],
            ["閉まる", "しまる", "shut/close"],
            ["閉める", "しめる", "shut/close"]
        ]],
        ["案", "アン / plan, suggestion, guide", [
            ["案内", "あんない", "information"],
            ["案", "あん", "a suggestion, proposal, plan"]
        ]],
        ["内", "ナイ / うち / inside, domestic", [
            ["家内", "かない", "wife"],
            ["内側", "うちがわ", "inside"],
            ["以内", "いない", "within ..."],
            ["国内", "こくない", "domestic"]
        ]],
        ["予", "ヨ / in advance, previous", [
            ["予定", "よてい", "a plan/schedule"],
            ["予習", "よしゅう", "preparation (pre-study)"]
        ]],
        ["約", "ヤク / promise, approximate", [
            ["予約", "よやく", "an appointment/reservation"],
            ["約～", "やく～", "approximately"]
        ]],

        # 第2週 2日目：禁煙
        ["煙", "エン / けむり / smoke", [
            ["禁煙", "きんえん", "no smoking"],
            ["煙", "けむり", "smoke"]
        ]],
        ["当", "トウ / あ-たる / hit, right, this", [
            ["本当", "ほんとう", "the truth"],
            ["当たる", "あたる", "hit/win"],
            ["当〇〇", "とう〇〇", "this ○○"],
            ["当たり前", "あたりまえ", "natural, not surprising"]
        ]],
        ["全", "ゼン / whole, entire, safe", [
            ["全部", "ぜんぶ", "all/whole"],
            ["全席", "ぜんせき", "all seats"],
            ["安全(な)", "あんぜん(な)", "safe"]
        ]],
        ["客", "キャク / guest, customer", [
            ["客", "きゃく", "a customer"],
            ["お客様", "おきゃくさま", "a customer"]
        ]],
        ["様", "ヨウ / さま / manner, honorific", [
            ["様子", "ようす", "appearance/situation"],
            ["〇〇様", "〇〇さま", "honorific added to people's names"]
        ]],
        ["解", "カイ / untie, solve, understand", [
            ["理解", "りかい", "understand"],
            ["解説", "かいせつ", "an explanation"],
            ["解答", "かいとう", "an answer"],
            ["分解", "ぶんかい", "take apart"]
        ]],
        ["協", "キョウ / cooperate", [
            ["協力", "きょうりょく", "cooperation"]
        ]],
        ["願", "ねが-う / wish, request", [
            ["願う", "ねがう", "wish"]
        ]],

        # 第2週 3日目：観光地図
        ["観", "カン / look at, view, sight", [
            ["観光", "かんこう", "sightseeing"],
            ["観客", "かんきゃく", "spectator, audience"]
        ]],
        ["園", "エン / garden, park", [
            ["動物園", "どうぶつえん", "a zoo"]
        ]],
        ["港", "コウ / みなと / harbor, port", [
            ["空港", "くうこう", "an airport"],
            ["港", "みなと", "a port"],
            ["〇〇港", "〇〇こう", "〇〇 port"]
        ]],
        ["遊", "ユウ / あそ-ぶ / play, travel", [
            ["遊園地", "ゆうえんち", "an amusement park"],
            ["遊ぶ", "あそぶ", "play"]
        ]],
        ["美", "ビ / うつく-しい / beauty, beautiful", [
            ["美術館", "びじゅつかん", "a museum"],
            ["美しい", "うつくしい", "beautiful"],
            ["美人", "びじん", "a beautiful woman"]
        ]],
        ["術", "ジュツ / art, technique, surgery", [
            ["美術", "びじゅつ", "(fine) art"],
            ["手術", "しゅじゅつ", "surgery, operation"],
            ["技術", "ぎじゅつ", "technique, technology"]
        ]],
        ["神", "シン / ジン / かみ / god, spirit, nerve", [
            ["神社", "じんじゃ", "a shrine"],
            ["神様", "かみさま", "a god"],
            ["神経質(な)", "しんけいしつ(な)", "nervous"]
        ]],
        ["寺", "ジ / てら / temple", [
            ["〇〇寺", "〇〇じ", "〇〇 Temple"],
            ["お寺", "おてら", "a temple"]
        ]],

        # 第2週 4日目：街の地図
        ["役", "ヤク / service, duty, office", [
            ["市役所", "しやくしょ", "the city office"],
            ["役員", "やくいん", "an officer/member"],
            ["役に立つ", "やくにたつ", "useful"]
        ]],
        ["郵", "ユウ / mail, post", [
            ["郵便", "ゆうびん", "mail"]
        ]],
        ["局", "キョク / bureau, office, department", [
            ["郵便局", "ゆうびんきょく", "a post office"],
            ["薬局", "やっきょく", "a drug store"]
        ]],
        ["交", "コウ / mix, association, traffic", [
            ["交番", "こうばん", "a police box"],
            ["交換", "こうかん", "exchange"],
            ["交通", "こうつう", "transport"]
        ]],
        ["差", "サ / さ-す / difference, submit", [
            ["差", "さ", "difference"],
            ["差し出す", "さしだす", "hand in"],
            ["差出人", "さしだしにん", "a sender"]
        ]],
        ["点", "テン / point, spot, score", [
            ["交差点", "こうさてん", "an intersection"],
            ["～点", "～てん", "... points"],
            ["点数", "てんすう", "score/points"]
        ]],
        ["橋", "キョウ / はし / bridge", [
            ["歩道橋", "ほどうきょう", "a foot bridge"],
            ["橋", "はし", "a bridge"]
        ]],
        ["公", "コウ / public, park", [
            ["公園", "こうえん", "a park"]
        ]],

        # 第2週 5日目：病院
        ["受", "ジュ / う-ける / receive, accept", [
            ["受信", "じゅしん", "receive"],
            ["受ける", "うける", "receive"],
            ["受験", "じゅけん", "an entrance examination"]
        ]],
        ["付", "つ-ける / つ-く / attach, stick to", [
            ["付ける", "つける", "put on"],
            ["受付", "うけつけ", "a reception"],
            ["片付ける", "かたづける", "tidy up"],
            ["付く", "つく", "stick to"]
        ]],
        ["科", "カ / department, course, science", [
            ["科学", "かがく", "science"],
            ["内科", "ないか", "internal medicine"],
            ["外科", "げか", "surgery department"],
            ["教科書", "きょうかしょ", "a textbook"]
        ]],
        ["鼻", "ビ / はな / nose", [
            ["耳鼻科", "じびか", "ears and nose department"],
            ["鼻", "はな", "a nose"]
        ]],
        ["婦", "フ / lady, woman", [
            ["婦人", "ふじん", "a woman"],
            ["主婦", "しゅふ", "a housewife"],
            ["産婦人科", "さんふじんか", "obstetrics and gynecology department"]
        ]],
        ["形", "ケイ / ギョウ / かたち / shape, form", [
            ["形式", "けいしき", "form / type"],
            ["整形外科", "せいけいげか", "orthopedics department"],
            ["人形", "にんぎょう", "a doll / puppet"],
            ["図形", "ずけい", "a diagram"],
            ["形", "かたち", "shape / form"]
        ]],
        ["骨", "コツ / ほね / bone, fracture", [
            ["骨折", "こっせつ", "a bone fracture"],
            ["骨", "ほね", "a bone"]
        ]],
        ["折", "セツ / お-る / お-れる / fold, break, turn", [
            ["右折", "うせつ", "a right turn"],
            ["折る", "おる", "break / fold (something)"],
            ["折れる", "おれる", "break"],
            ["左折", "させつ", "a left turn"],
            ["折り紙", "おりがみ", "folding paper"]
        ]],

        # 第2週 6日目：困ったときは
        ["困", "こま-る / trouble, distressed", [
            ["困る", "こまる", "be in trouble"]
        ]],
        ["消", "ショウ / き-える / け-す / extinguish, disappear", [
            ["消防", "しょうぼう", "fire fighting"],
            ["消す", "けす", "extinguish"],
            ["消える", "きえる", "be extinguished, disappear"],
            ["消しゴム", "けしゴム", "an eraser"]
        ]],
        ["防", "ボウ / ふせ-ぐ / prevent, defend", [
            ["予防", "よぼう", "precaution, prevention"],
            ["防ぐ", "ふせぐ", "prevent"]
        ]],
        ["救", "キュウ / すく-う / rescue, save", [
            ["救急車", "きゅうきゅうしゃ", "an ambulance"],
            ["救う", "すくう", "save"]
        ]],
        ["警", "ケイ / police, guard, caution", [
            ["警官", "けいかん", "a police officer"],
            ["警察", "けいさつ", "police"],
            ["警察署", "けいさつしょ", "a police station"]
        ]],
        ["察", "サツ / inspect, perceive", [
            ["警察", "けいさつ", "police"],
            ["警察署", "けいさつしょ", "a police station"]
        ]],
        ["故", "コ / accident, reason, late", [
            ["事故", "じこ", "an accident"],
            ["故〇〇", "こ〇〇", "the late Mr."],
            ["故障", "こしょう", "break-down, failure"]
        ]],
        ["伝", "デン / つた-える / convey, help", [
            ["伝言", "でんごん", "a message"],
            ["伝える", "つたえる", "convey / communicate"],
            ["手伝う", "てつだう", "help"]
        ]],

        # 第2週 7日目：クイズ②
        ["黄", "き / yellow", [
            ["黄色", "きいろ", "yellow"],
            ["黄色い", "きいろい", "yellow"]
        ]],
        ["絵", "カイ / エ / picture, painting", [
            ["絵画", "かいが", "a painting"],
            ["絵", "え", "a picture"],
            ["絵本", "えほん", "a picture book"]
        ]],
        ["組", "くみ / く-む / pair, team up, program", [
            ["～組", "～くみ", "pairs"],
            ["組む", "くむ", "team up"],
            ["番組", "ばんぐみ", "a (TV) program"],
            ["組み立てる", "くみたてる", "put together"]
        ]],
        ["束", "ソク / たば / bundle, promise", [
            ["約束", "やくそく", "a promise / appointment"],
            ["～束", "～たば", "a bundle, a bunch"],
            ["花束", "はなたば", "bunch of flowers"]
        ]],
        ["授", "ジュ / impart, instruct, class", [
            ["授業", "じゅぎょう", "a class"],
            ["教授", "きょうじゅ", "a professor"]
        ]],
        ["渡", "わた-る / わた-す / cross over, hand over", [
            ["渡る", "わたる", "cross over"],
            ["渡す", "わたす", "hand over"]
        ]],
        ["昔", "むかし / long ago, past", [
            ["昔", "むかし", "long ago, past"]
        ]]
    ]
}

# Chapter 3 (Week 3, 7 Days): Kanji data
CHAPTER_3_DATA = {
    "Chapter 3 (Week 3)": [
        # 第3週 1日目：要冷蔵
        ["要", "ヨウ / い-る / necessary, need, important", [
            ["必要(な)", "ひつよう(な)", "necessary"],
            ["要る", "いる", "need"],
            ["重要(な)", "じゅうよう(な)", "important"]
        ]],
        ["冷", "レイ / つめ-たい / ひ-やす / ひ-える / さ-める / さ-ます / cool, cold", [
            ["冷房", "れいぼう", "air-conditioning"],
            ["冷やす", "ひやす", "cool"],
            ["冷める", "さめる", "cool down"],
            ["冷たい", "つめたい", "cold"],
            ["冷える", "ひえる", "become cold"],
            ["冷ます", "さます", "cool something"]
        ]],
        ["蔵", "ゾウ / storehouse, hide", [
            ["冷蔵庫", "れいぞうこ", "a refrigerator"]
        ]],
        ["凍", "トウ / こお-る / freeze", [
            ["冷凍庫", "れいとうこ", "a freezer"],
            ["凍る", "こおる", "freeze"]
        ]],
        ["庫", "コ / warehouse, storehouse", [
            ["金庫", "きんこ", "a safe"],
            ["車庫", "しゃこ", "a garage"]
        ]],
        ["召", "め-す / call, invite, eat/drink", [
            ["召し上がる", "めしあがる", "eat (polite form)"]
        ]],
        ["保", "ホ / preserve, protect, keep", [
            ["保存する", "ほぞんする", "preserve"]
        ]],
        ["存", "ゾン / exist, know", [
            ["ご存じです", "ごぞんじです", "know (humble form)"],
            ["存じません", "ぞんじません", "I do not know"]
        ]],
        ["必", "ヒツ / かなら-ず / certain, necessary", [
            ["必要(な)", "ひつよう(な)", "necessary"],
            ["必ず", "かならず", "always / certainly"],
            ["必死(に)", "ひっし(に)", "desperate(ly)"]
        ]],

        # 第3週 2日目：消費期限
        ["費", "ヒ / expense, cost", [
            ["費用", "ひよう", "an expense, cost"],
            ["旅費", "りょひ", "travelling expenses"],
            ["消費者", "しょうひしゃ", "consumers"],
            ["会費", "かいひ", "a membership fee"]
        ]],
        ["期", "キ / period, time", [
            ["期間", "きかん", "a period of time"],
            ["長期", "ちょうき", "a long period"],
            ["定期券", "ていきけん", "a commuter pass"],
            ["短期", "たんき", "a short period"]
        ]],
        ["限", "ゲン / かぎ-る / limit, restrict", [
            ["期限", "きげん", "a time limit"],
            ["限る", "かぎる", "limit, restrict"],
            ["限度", "げんど", "a limit"],
            ["限定", "げんてい", "limitation"]
        ]],
        ["製", "セイ / make, manufacture", [
            ["○○製", "○○せい", "made in/of ○○"],
            ["製品", "せいひん", "a product"]
        ]],
        ["造", "ゾウ / つく-る / create, make", [
            ["製造", "せいぞう", "manufacture"],
            ["造る", "つくる", "make"]
        ]],
        ["賞", "ショウ / prize, reward", [
            ["賞", "しょう", "a prize"],
            ["賞金", "しょうきん", "prize money"],
            ["賞味期限", "しょうみきげん", "Best before"],
            ["賞品", "しょうひん", "a prize"]
        ]],
        ["法", "ホウ / method, law", [
            ["方法", "ほうほう", "a method / way"],
            ["文法", "ぶんぽう", "grammar"]
        ]],
        ["温", "オン / あたた-かい / warm, temperature", [
            ["温度", "おんど", "temperature"],
            ["温かい", "あたたかい", "warm"],
            ["気温", "きおん", "temperature"],
            ["常温", "じょうおん", "normal temperature"]
        ]],

        # 第3週 3日目：自動販売機
        ["販", "ハン / sell, marketing", [
            ["販売", "はんばい", "sell"],
            ["自動販売機", "じどうはんばいき", "a vending machine"]
        ]],
        ["機", "キ / machine, opportunity", [
            ["飛行機", "ひこうき", "an airplane"],
            ["機会", "きかい", "an opportunity"],
            ["機械", "きかい", "a machine"]
        ]],
        ["増", "ゾウ / ふ-える / ふ-やす / increase, add", [
            ["増加", "ぞうか", "an increase"],
            ["増える", "ふえる", "increase"],
            ["増やす", "ふやす", "increase, add"]
        ]],
        ["減", "ゲン / へ-る / へ-らす / decrease, reduce", [
            ["減少", "げんしょう", "a decrease"],
            ["減る", "へる", "decrease, reduce"],
            ["減らす", "へらす", "decrease"]
        ]],
        ["量", "リョウ / quantity, amount", [
            ["量", "りょう", "quantity"],
            ["増量", "ぞうりょう", "increase the amount"],
            ["数量", "すうりょう", "amount"],
            ["減量", "げんりょう", "a loss in quantity (weight)"]
        ]],
        ["氷", "こおり / ice", [
            ["氷", "こおり", "ice"]
        ]],
        ["返", "ヘン / かえ-す / return, reply", [
            ["返事", "へんじ", "a reply"],
            ["返す", "かえす", "return (something)"],
            ["返却", "へんきゃく", "return"]
        ]],
        ["湯", "ゆ / hot water", [
            ["(お)湯", "(お)ゆ", "hot water"]
        ]],

        # 第3週 4日目：レシピ
        ["材", "ザイ / ingredients, materials", [
            ["材料", "ざいりょう", "ingredients, materials"],
            ["教材", "きょうざい", "teaching material"]
        ]],
        ["卵", "たまご / egg", [
            ["卵", "たまご", "an egg"],
            ["卵焼き", "たまごやき", "a Japanese omlet"]
        ]],
        ["乳", "ニュウ / milk", [
            ["牛乳", "ぎゅうにゅう", "milk"]
        ]],
        ["粉", "こな / こ / powder, flour", [
            ["粉", "こな", "powder, flour"],
            ["小麦粉", "こむぎこ", "wheat flour"]
        ]],
        ["袋", "ふくろ / bag", [
            ["袋", "ふくろ", "a bag"],
            ["ごみ袋", "ごみぶくろ", "a garbage bag"],
            ["足袋", "たび", "tabi (traditional Japanese socks worn with a kimono)"],
            ["紙袋", "かみぶくろ", "a paper bag"],
            ["手袋", "てぶくろ", "gloves"]
        ]],
        ["混", "コン / ま-ぜる / mix, blend", [
            ["混雑", "こんざつ", "congestion"],
            ["混ぜる", "まぜる", "be mixed"]
        ]],
        ["焼", "や-く / や-ける / roast, grill, bake", [
            ["焼く", "やく", "roast, grill"],
            ["焼ける", "やける", "be burnt/baked"]
        ]],
        ["表", "ヒョウ / おもて / あらわ-す / surface, table, express", [
            ["表", "ひょう", "a table (in written documents)"],
            ["発表", "はっぴょう", "an announcement"],
            ["表", "おもて", "a surface, face"],
            ["表面", "ひょうめん", "a surface"],
            ["代表", "だいひょう", "a representative"],
            ["表す", "あらわす", "show, express"]
        ]],
        ["裏", "うら / reverse, back", [
            ["裏", "うら", "reverse, back"],
            ["裏返す", "うらがえす", "turnover"]
        ]],

        # 第3週 5日目：コピー機・留守番電話
        ["留", "リュウ / ル / と-める / detain, stop, stay", [
            ["留学", "りゅうがく", "study abroad"],
            ["保留", "ほりゅう", "reservation / suspension"],
            ["留守番", "るすばん", "stay at home"],
            ["書留", "かきとめ", "registration (registered mail)"]
        ]],
        ["守", "シュ / ス / まも-る / defence, protect", [
            ["守備", "しゅび", "defence"],
            ["守る", "まもる", "protect"],
            ["留守", "るす", "absence"]
        ]],
        ["濃", "こ-い / concentrate, dark", [
            ["濃い", "こい", "concentrate, dark (color)"]
        ]],
        ["薄", "うす-い / thin, light, weak", [
            ["薄い", "うすい", "thin (material), light (color), weak (drink)"]
        ]],
        ["部", "ブ / part, section, department", [
            ["部分", "ぶぶん", "part"],
            ["部長", "ぶちょう", "department head / manager"],
            ["学部", "がくぶ", "faculty"],
            ["部屋", "へや", "a room"]
        ]],
        ["数", "スウ / かず / かぞ-える / number, count", [
            ["数字", "すうじ", "number"],
            ["数", "かず", "a number"],
            ["数学", "すうがく", "mathematics"],
            ["数える", "かぞえる", "count"]
        ]],
        ["件", "ケン / subject, matter, incident", [
            ["件名", "けんめい", "subject"],
            ["用件", "ようけん", "a business/matter"],
            ["事件", "じけん", "an incident"]
        ]],
        ["再", "サイ / サ / re-, again", [
            ["再入国", "さいにゅうこく", "re-enter a country"],
            ["再生", "さいせい", "regenerate, recycle"],
            ["再ダイヤル", "さいダイヤル", "redial"],
            ["再来週", "さらいしゅう", "the week after next"]
        ]],

        # 第3週 6日目：携帯電話
        ["接", "セツ / touch, connect", [
            ["接続", "せつぞく", "connect"],
            ["面接", "めんせつ", "interview"]
        ]],
        ["続", "ゾク / つづ-く / つづ-ける / continue", [
            ["接続", "せつぞく", "connect"],
            ["続く", "つづく", "continue"],
            ["続ける", "つづける", "continue"]
        ]],
        ["示", "ジ / しめ-す / show, indicate", [
            ["表示", "ひょうじ", "indication, expression"],
            ["示す", "しめす", "show, point out"],
            ["指示", "しじ", "a direction, instruction"]
        ]],
        ["戻", "もど-る / もど-す / return", [
            ["戻る", "もどる", "return"],
            ["戻す", "もどす", "return, put back"]
        ]],
        ["完", "カン / completion, perfect", [
            ["完了", "かんりょう", "completion"],
            ["完全(な)", "かんぜん(な)", "perfect, complete"]
        ]],
        ["了", "リョウ / finish, understand", [
            ["了解", "りょうかい", "understand, agree"],
            ["終了", "しゅうりょう", "end, expiration"]
        ]],
        ["登", "トウ / ト / のぼ-る / registration, climb", [
            ["登録", "とうろく", "registration"],
            ["登る", "のぼる", "climb"],
            ["登山", "とざん", "mountain climbing"]
        ]],
        ["録", "ロク / record", [
            ["記録", "きろく", "a record"],
            ["録音", "ろくおん", "record (sound)"],
            ["録画", "ろくが", "record (video)"]
        ]],

        # 第3週 7日目：クイズ③ どれが入る？
        ["育", "イク / そだ-つ / そだ-てる / education, raise, grow", [
            ["教育", "きょういく", "education"],
            ["育てる", "そだてる", "raise, bring up"],
            ["育つ", "そだつ", "grow"]
        ]],
        ["種", "シュ / たね / type, seed", [
            ["種類", "しゅるい", "type, kind"],
            ["種", "たね", "a seed"]
        ]],
        ["類", "ルイ / document, classification, race", [
            ["書類", "しょるい", "a document"],
            ["分類", "ぶんるい", "classification"],
            ["人類", "じんるい", "the human race"]
        ]],
        ["師", "シ / teacher, nurse, doctor", [
            ["教師", "きょうし", "a teacher"],
            ["看護師", "かんごし", "a nurse"],
            ["医師", "いし", "a doctor"]
        ]],
        ["妻", "サイ / つま / wife", [
            ["夫妻", "ふさい", "husband and wife"],
            ["妻", "つま", "wife"]
        ]],
        ["馬", "バ / うま / horse", [
            ["馬", "うま", "a horse"],
            ["乗馬", "じょうば", "horse riding"]
        ]],
        ["石", "セキ / セツ / いし / soap, stone, oil", [
            ["石けん", "せっけん", "soap"],
            ["石", "いし", "stone"],
            ["石油", "せきゆ", "oil (petroleum)"]
        ]]
    ]
}

def generate_soma_document(subset_dict, filename="Kanji_Practice_Chapter3_5cols.docx"):
    doc = Document()
    
    # Configure margins for extra printable width (7.7 inches)
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(0.4)
        s.bottom_margin = Inches(0.4)
        s.left_margin = Inches(0.4)
        s.right_margin = Inches(0.4)

    # 1 WORD column (1.7 in) + 5 Practice columns (1.2 in each) = 7.7 in total printable width
    COL_WIDTHS = [Inches(1.7)] + [Inches(1.2)] * 5
    
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

            # Practice Grid Table for this Kanji (6 columns total: WORD + 5 Practice columns)
            table = doc.add_table(rows=1, cols=6)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(table, color="D3C7B6")

            # Header Row
            hdr_row = table.rows[0]
            hdr_titles = ["WORD", "1", "2", "3", "4", "5"]
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

            # Add 3 practice sub-rows per vocabulary word block
            for w_kanji, w_hira, w_eng in vocab_list:
                r0 = table.add_row()
                r1 = table.add_row()
                r2 = table.add_row()

                for r in [r0, r1, r2]:
                    trPr = r._element.get_or_add_trPr()
                    trHeight = parse_xml(f'<w:trHeight {nsdecls("w")} w:val="600" w:hRule="atLeast"/>')
                    trPr.append(trHeight)
                    for i, c in enumerate(r.cells):
                        c.width = COL_WIDTHS[i]
                        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                # Merge column 0 across r0, r1, r2 (3 rows block)
                info_cell = r0.cells[0]
                info_cell.merge(r1.cells[0])
                info_cell.merge(r2.cells[0])
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
    generate_soma_document(CHAPTER_3_DATA, "Kanji_Practice_Chapter3_5cols.docx")

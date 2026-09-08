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

# Chapter 4 (Week 4, 7 Days): Kanji data
CHAPTER_4_DATA = {
    "Chapter 4 (Week 4)": [
        # 第4週 1日目：日用品
        ["砂", "サ / すな / sand, sugar", [
            ["砂糖", "さとう", "sugar"],
            ["砂", "すな", "sand"]
        ]],
        ["塩", "エン / しお / salt", [
            ["食塩", "しょくえん", "table salt"],
            ["塩", "しお", "salt"]
        ]],
        ["油", "ユ / あぶら / oil", [
            ["しょう油", "しょうゆ", "soy sauce"],
            ["灯油", "とうゆ", "kerosene"],
            ["石油", "せきゆ", "oil (petroleum)"],
            ["油", "あぶら", "oil"]
        ]],
        ["緑", "リョク / みどり / green, green tea", [
            ["緑茶", "りょくちゃ", "green tea"],
            ["緑(色)", "みどり(いろ)", "green"]
        ]],
        ["紅", "コウ / べに / crimson, tea, lipstick", [
            ["紅茶", "こうちゃ", "tea"],
            ["口紅", "くちべに", "lipstick"]
        ]],
        ["冊", "サツ / counter for books", [
            ["～冊", "～さつ", "counter for books"],
            ["冊数", "さっすう", "the number of copies"]
        ]],
        ["個", "コ / counter of general objects, individual", [
            ["～個", "～こ", "counter of general objects"],
            ["個人", "こじん", "individual (person)"],
            ["個数", "こすう", "the number of items"],
            ["団体", "だんたい", "a group"]
        ]],
        ["枚", "マイ / counter for flat objects", [
            ["～枚", "～まい", "counter for flat objects"],
            ["枚数", "まいすう", "the number of sheets/copies"]
        ]],

        # 第4週 2日目：広告メール
        ["告", "コク / announce, advertise", [
            ["広告", "こうこく", "an advertisement"]
        ]],
        ["利", "リ / profit, convenient, use", [
            ["便利(な)", "べんり(な)", "convenient"],
            ["利用", "りよう", "use"]
        ]],
        ["割", "わ-る / わ-れる / break, ratio, discount", [
            ["割る", "わる", "break"],
            ["割れる", "われる", "crack/cleave"],
            ["割合", "わりあい", "a ratio, a percentage"],
            ["割引", "わりびき", "a discount"]
        ]],
        ["倍", "バイ / times, double", [
            ["～倍", "～ばい", "... times (quantity - e.g. twice as many)"],
            ["倍", "ばい", "2倍"]
        ]],
        ["値", "ね / price", [
            ["値段", "ねだん", "a price"],
            ["値上げ", "ねあげ", "a price increase"],
            ["値下げ", "ねさげ", "a price reduction"]
        ]],
        ["商", "ショウ / trade, goods, shop", [
            ["商品", "しょうひん", "goods"],
            ["商店", "しょうてん", "a shop"]
        ]],
        ["支", "シ / branch, support, pay", [
            ["支店", "してん", "a branch"],
            ["支社", "ししゃ", "a branch office"],
            ["本店", "ほんてん", "the main branch of a store"],
            ["本社", "ほんしゃ", "the head office"]
        ]],
        ["払", "はら-う / pay", [
            ["払う", "はらう", "pay"],
            ["支払い", "しはらい", "a payment"]
        ]],

        # 第4週 3日目：通信販売
        ["米", "ベイ / こめ / rice, USA", [
            ["米国", "べいこく", "the United States of America"],
            ["米", "こめ", "rice"]
        ]],
        ["級", "キュウ / class, grade", [
            ["高級", "こうきゅう", "high class/grade"],
            ["上級", "じょうきゅう", "advanced level"],
            ["中級", "ちゅうきゅう", "intermediate level"]
        ]],
        ["残", "ザン / のこ-る / のこ-す / remain, leave", [
            ["残業", "ざんぎょう", "overtime work"],
            ["残る", "のこる", "remain"],
            ["残り", "のこり", "the remainder"],
            ["残す", "のこす", "leave, leave behind"]
        ]],
        ["型", "かた / model, type", [
            ["大型", "おおがた", "large, jumbo"],
            ["新型", "しんがた", "new-model"],
            ["小型", "こがた", "small-sized"]
        ]],
        ["税", "ゼイ / tax", [
            ["消費税", "しょうひぜい", "consumption tax"],
            ["税金", "ぜいきん", "a tax"]
        ]],
        ["込", "こ-む / crowd, include", [
            ["込む", "こむ", "congest/crowd"],
            ["振り込む", "ふりこむ", "transfer money to a person's account"],
            ["税込", "ぜいこみ", "tax included"]
        ]],
        ["価", "カ / price, value", [
            ["価格", "かかく", "a price"],
            ["定価", "ていか", "a fixed price"]
        ]],
        ["格", "カク / status, price, pass", [
            ["合格", "ごうかく", "pass an exam"],
            ["格安", "かくやす", "a bargain"]
        ]],

        # 第4週 4日目：申込書
        ["申", "シン / もう-す / apply, say (humble)", [
            ["申し込む", "もうしこむ", "apply"],
            ["申告", "しんこく", "a declaration"],
            ["申す", "もうす", "say (humble form)"],
            ["申込書", "もうしこみしょ", "an application form"],
            ["申請", "しんせい", "application"],
            ["申し上げる", "もうしあげる", "say (very humble form)"]
        ]],
        ["記", "キ / record, sign, diary", [
            ["記入", "きにゅう", "entry"],
            ["記号", "きごう", "a sign, a symbol"],
            ["日記", "にっき", "a diary / journal"],
            ["記事", "きじ", "an article"]
        ]],
        ["例", "レイ / たと-える / example", [
            ["例", "れい", "an example"],
            ["例えば", "たとえば", "for example"]
        ]],
        ["齢", "レイ / age", [
            ["年齢", "ねんれい", "age"],
            ["高齢", "こうれい", "old age"]
        ]],
        ["歳", "サイ / years old", [
            ["～歳", "～さい", "... years old"],
            ["二十歳", "にじゅっさい / はたち", "20 years old"]
        ]],
        ["性", "セイ / gender, personality", [
            ["性別", "せいべつ", "sex/gender"],
            ["女性", "じょせい", "woman"],
            ["性格", "せいかく", "personality"],
            ["男性", "だんせい", "man"]
        ]],
        ["連", "レン / つ-れる / connect, take along", [
            ["連休", "れんきゅう", "a holiday"],
            ["連れて行く", "つれていく", "take someone to ..."],
            ["連れて来る", "つれてくる", "bring someone to ..."]
        ]],
        ["絡", "ラク / contact", [
            ["連絡", "れんらく", "contact/connection"]
        ]],

        # 第4週 5日目：注文
        ["届", "とど-ける / とど-く / deliver, arrive", [
            ["届ける", "とどける", "deliver"],
            ["届く", "とどく", "arrive (mail)"]
        ]],
        ["宅", "タク / house, home", [
            ["自宅", "じたく", "one's house / home"],
            ["宅配", "たくはい", "deliver to someone's house"],
            ["お宅", "おたく", "house/home (respectful form)"]
        ]],
        ["配", "ハイ / くば-る / delivery, distribute", [
            ["配達", "はいたつ", "delivery"],
            ["心配", "しんぱい", "anxiety/worry"],
            ["配送料", "はいそうりょう", "a shipping charge"],
            ["配る", "くばる", "distribute"]
        ]],
        ["希", "キ / hope", [
            ["希望", "きぼう", "hope"]
        ]],
        ["望", "ボウ / のぞ-む / hope, want", [
            ["失望", "しつぼう", "despair"],
            ["望む", "のぞむ", "want, hope for"]
        ]],
        ["荷", "カ / luggage, load", [
            ["入荷", "にゅうか", "receipt (of goods)"],
            ["荷物", "にもつ", "luggage"],
            ["出荷", "しゅっか", "shipment"],
            ["手荷物", "てにもつ", "hand luggage"]
        ]],
        ["換", "カン / か-える / exchange", [
            ["交換", "こうかん", "exchange"],
            ["乗り換え", "のりかえ", "changing trains"],
            ["代金引換", "だいきんひきかえ", "cash on delivery"]
        ]],
        ["額", "ガク / amount, sum", [
            ["金額", "きんがく", "an amount/sum (of money)"],
            ["半額", "はんがく", "half price"]
        ]],

        # 第4週 6日目：不在通知
        ["在", "ザイ / present, exist", [
            ["不在", "ふざい", "absence"],
            ["現在", "げんざい", "present (time)"]
        ]],
        ["取", "と-る / take, receive", [
            ["取る", "とる", "take"],
            ["受け取る", "うけとる", "receive, take"],
            ["受取人", "うけとりにん", "a recipient"]
        ]],
        ["預", "ヨ / あず-ける / deposit, entrust", [
            ["預金", "よきん", "a money deposit"],
            ["預ける", "あずける", "entrust"]
        ]],
        ["衣", "イ / clothing", [
            ["衣類", "いるい", "clothing"],
            ["衣服", "いふく", "clothes"]
        ]],
        ["参", "サン / まい-る / participate, go/come (humble)", [
            ["参加", "さんか", "participate"],
            ["参る", "まいる", "go / come (humble form)"],
            ["参考書", "さんこうしょ", "a reference book"]
        ]],
        ["達", "タツ / progress, friend, special delivery", [
            ["上達", "じょうたつ", "make progress"],
            ["友達", "ともだち", "a friend"],
            ["速達", "そくたつ", "a special delivery"]
        ]],
        ["勤", "キン / つと-める / work, employ", [
            ["通勤", "つうきん", "commuting to work"],
            ["勤める", "つとめる", "work / be employed"]
        ]],
        ["帯", "タイ / おび / phone, belt, time slot", [
            ["携帯(電話)", "けいたい(でんわ)", "a cell phone / mobile phone"],
            ["帯", "おび", "a belt/sash"],
            ["時間帯", "じかんたい", "a time zone, a time slot"]
        ]],

        # 第4週 7日目：クイズ④ 読みはどちら？
        ["細", "ほそ-い / こま-かい / fine, thin, small", [
            ["細い", "ほそい", "fine/thin"],
            ["細かい", "こまかい", "fine/small"]
        ]],
        ["戸", "と / door", [
            ["戸", "と", "a door"],
            ["戸だな", "とだな", "a closet, a cupboard"],
            ["雨戸", "あまど", "a sliding storm door"]
        ]],
        ["湖", "コ / みずうみ / lake", [
            ["びわ湖", "びわこ", "Lake Biwa"],
            ["湖", "みずうみ", "a lake"]
        ]],
        ["船", "セン / ふね / ふな / ship, boat", [
            ["風船", "ふうせん", "a balloon"],
            ["船", "ふね", "a boat, ship"],
            ["船長", "せんちょう", "a captain"],
            ["船便", "ふなびん", "surface/sea mail"]
        ]],
        ["角", "カク / かど / angle, corner, square", [
            ["角度", "かくど", "an angle"],
            ["四角い", "しかくい", "square"],
            ["三角形", "さんかくけい", "a triangle"],
            ["角", "かど", "a corner"]
        ]],
        ["夫", "フ / フウ / おっと / husband, couple", [
            ["夫妻", "ふさい", "husband and wife"],
            ["夫婦", "ふうふ", "a married couple"],
            ["夫", "おっと", "a husband"]
        ]],
        ["苦", "ク / くる-しい / にが-い / bitter, trying", [
            ["苦労", "くろう", "distressful, trying"],
            ["苦い", "にがい", "bitter"],
            ["苦しい", "くるしい", "distressful, trying"],
            ["苦手(な)", "にがて(な)", "a weak point"]
        ]]
    ]
}

# Chapter 5 (Week 5, 7 Days): Kanji data
CHAPTER_5_DATA = {
    "Chapter 5 (Week 5)": [
        # 第5週 1日目：メールを送る
        ["礼", "レイ / thanks", [
            ["お礼", "おれい", "thanks"]
        ]],
        ["伺", "うかが-う / visit, ask (humble form)", [
            ["伺う", "うかがう", "visit, ask (humble form)"]
        ]],
        ["遅", "チ / おそ-い / おく-れる / tardiness, slow, be late", [
            ["遅刻", "ちこく", "tardiness"],
            ["遅い", "おそい", "slow"],
            ["遅れる", "おくれる", "be late"]
        ]],
        ["失", "シツ / rudeness, failure, mistake, be excused", [
            ["失礼(な)", "しつれい(な)", "rudeness"],
            ["失敗", "しっぱい", "failure, mistake"],
            ["失礼する", "しつれいする", "be excused"]
        ]],
        ["汗", "あせ / perspiration, sweat", [
            ["汗", "あせ", "perspiration, sweat"],
            ["汗をかく", "あせをかく", "sweat"]
        ]],
        ["念", "ネン / regret, disappointment, commemoration", [
            ["残念(な)", "ざんねん(な)", "regret, disappointment"],
            ["記念", "きねん", "commemoration"]
        ]],
        ["涙", "なみだ / tear(s), weep", [
            ["涙", "なみだ", "tear(s)"],
            ["涙を流す", "なみだをながす", "weep"]
        ]],
        ["笑", "わら-う / え-む / laugh, smile, laughter", [
            ["笑う", "わらう", "laugh / smile"],
            ["笑顔", "えがお", "smile / smiling face"],
            ["笑い", "わらい", "laughter"]
        ]],

        # 第5週 2日目：アンケート
        ["調", "チョウ / しら-べる / condition, look up, emphasis", [
            ["調子", "ちょうし", "condition"],
            ["調べる", "しらべる", "look up something, investigate"],
            ["強調", "きょうちょう", "emphasis, stress"]
        ]],
        ["査", "サ / investigation, inquiry", [
            ["調査", "ちょうさ", "an investigation/inquiry"]
        ]],
        ["移", "イ / うつ-る / うつ-す / move, transfer, shift", [
            ["移動", "いどう", "move / transfer"],
            ["移る", "うつる", "move / shift"],
            ["移す", "うつす", "move / shift (something)"]
        ]],
        ["難", "ナン / むずか-しい / difficulty, difficult, problem", [
            ["困難(な)", "こんなん(な)", "difficulty"],
            ["難しい", "むずかしい", "difficult"],
            ["難問", "なんもん", "a difficult problem"]
        ]],
        ["簡", "カン / easy, simple", [
            ["簡単(な)", "かんたん(な)", "easy"]
        ]],
        ["単", "タン / vocabulary, unit, credit", [
            ["単語", "たんご", "vocabulary"],
            ["単位", "たんい", "a unit, credit"]
        ]],
        ["感", "カン / feel, inspiration", [
            ["感じる", "かんじる", "feel"],
            ["感動", "かんどう", "inspiration"]
        ]],
        ["想", "ソウ / impressions, thoughts, forecast", [
            ["感想", "かんそう", "impressions, thoughts"],
            ["予想", "よそう", "anticipation, forecast"]
        ]],

        # 第5週 3日目：日本語クラス
        ["練", "レン / practice", [
            ["練習", "れんしゅう", "practice"]
        ]],
        ["最", "サイ / もっと-も / recently, last, beginning, most", [
            ["最近", "さいきん", "recently"],
            ["最後", "さいご", "last / end"],
            ["最初", "さいしょ", "first / beginning"],
            ["最も", "もっとも", "most"]
        ]],
        ["適", "テキ / correct, appropriate, comfortable", [
            ["適当(な)", "てきとう(な)", "correct / appropriate"],
            ["快適(な)", "かいてき(な)", "comfortable"]
        ]],
        ["選", "セン / えら-ぶ / election, choose, player, athlete", [
            ["選挙", "せんきょ", "an election"],
            ["選ぶ", "えらぶ", "choose"],
            ["選手", "せんしゅ", "a player, an athlete"]
        ]],
        ["違", "ちが-う / ちが-える / different, wrong, mistake", [
            ["違う", "ちがう", "different, wrong"],
            ["間違い", "まちがい", "a mistake"],
            ["間違う", "まちがう", "make a mistake"],
            ["間違える", "まちがえる", "make a mistake"]
        ]],
        ["直", "チョク / なお-る / なお-す / straight line, repair, direct, fix", [
            ["直線", "ちょくせん", "a straight line"],
            ["直る", "なおる", "be repaired"],
            ["直接", "ちょくせつ", "direct"],
            ["直す", "なおす", "fix"]
        ]],
        ["復", "フク / review, recovery, round trip", [
            ["復習", "ふくしゅう", "review"],
            ["回復", "かいふく", "recovery/recuperation"],
            ["往復", "おうふく", "a round trip"]
        ]],
        ["辞", "ジ / や-める / dictionary, resign, retire", [
            ["辞書", "じしょ", "a dictionary"],
            ["辞める", "やめる", "resign, retire"]
        ]],
        ["宿", "シュク / やど / homework, lodgings, inn, hotel", [
            ["宿題", "しゅくだい", "a homework"],
            ["下宿", "げしゅく", "lodgings"],
            ["宿", "やど", "an inn, a hotel"]
        ]],

        # 第5週 4日目：作文
        ["昨", "サク / yesterday, last night, last year", [
            ["昨日", "さくじつ / きのう", "yesterday"],
            ["昨夜", "さくや", "last night"],
            ["昨年", "さくねん", "last year"]
        ]],
        ["君", "クン / きみ / honorific, you", [
            ["○○君", "○○くん", "honorific appended to names of males younger than oneself"],
            ["君", "きみ", "you"]
        ]],
        ["結", "ケツ / むす-ぶ / splendid, after all, tie, connect, conclude", [
            ["結構(な)", "けっこう(な)", "splendid, nice"],
            ["結局", "けっきょく", "after all"],
            ["「いいえ、結構です」", "「いいえ、けっこうです」", "No, thank you."],
            ["結ぶ", "むすぶ", "tie / connect / conclude"]
        ]],
        ["婚", "コン / marriage, engagement, honeymoon", [
            ["結婚", "けっこん", "a marriage"],
            ["婚約", "こんやく", "an engagement"],
            ["新婚旅行", "しんこんりょこう", "a honeymoon"]
        ]],
        ["祝", "シュク / いわ-う / holiday, celebrate, congratulate", [
            ["祝日", "しゅくじつ", "a holiday / festival day"],
            ["祝う", "いわう", "celebrate/congratulate"],
            ["お祝い", "おいわい", "celebration/congratulation"]
        ]],
        ["曲", "キョク / ま-がる / ま-げる / music piece, bend, curve", [
            ["曲", "きょく", "a piece of music"],
            ["曲がる", "まがる", "bend, turn a corner"],
            ["曲線", "きょくせん", "a curve"],
            ["曲げる", "まげる", "bend"]
        ]],
        ["奥", "おく / wife, inmost", [
            ["奥さん", "おくさん", "another person's wife"],
            ["奥", "おく", "inmost"]
        ]],
        ["寝", "ね-る / sleep, nap, late riser", [
            ["寝る", "ねる", "sleep"],
            ["昼寝", "ひるね", "a nap"],
            ["寝坊", "ねぼう", "late riser, sleepyhead"]
        ]],

        # 第5週 5日目：問診票―歯科で
        ["痛", "ツウ / いた-い / headache, sore, painful, stomach ache", [
            ["頭痛", "ずつう", "a headache"],
            ["痛い", "いたい", "sore, painful"],
            ["腹痛", "ふくつう", "stomach ache"]
        ]],
        ["熱", "ネツ / あつ-い / heat, fever, hot, enthusiasm", [
            ["熱", "ねつ", "heat, fever"],
            ["熱い", "あつい", "hot"],
            ["熱心(な)", "ねっしん(な)", "enthusiasm, zeal"]
        ]],
        ["虫", "むし / insect, decayed tooth", [
            ["虫", "むし", "an insect"]
        ]],
        ["歯", "シ / は / dentistry, dentist, teeth, decayed tooth", [
            ["歯科", "しか", "dentistry"],
            ["歯医者", "はいしゃ", "a dentist"],
            ["歯", "は", "teeth"],
            ["虫歯", "むしば", "a decayed tooth"]
        ]],
        ["治", "ジ / チ / なお-る / なお-す / politics, heal, treatment, cure", [
            ["政治", "せいじ", "politics, government"],
            ["治る", "なおる", "heal"],
            ["治療", "ちりょう", "a treatment"],
            ["治す", "なおす", "cure"]
        ]],
        ["汚", "きたな-い / よご-れる / dirty, become dirty", [
            ["汚い", "きたない", "dirty"],
            ["汚れる", "よごれる", "become dirty"]
        ]],
        ["並", "なら-ぶ / なら-べる / stand in line, alignment", [
            ["並ぶ", "ならぶ", "stand in a line"],
            ["歯並び", "はならび", "the alignment of your teeth"],
            ["並べる", "ならべる", "line up, set up"]
        ]],
        ["他", "タ / ほか / other", [
            ["他の人", "ほかのひと", "other"],
            ["その他", "そのた", "other"]
        ]],

        # 第5週 6日目：問診票―健康診断
        ["身", "シン / み / height, status, single, sashimi", [
            ["身長", "しんちょう", "height"],
            ["身分", "みぶん", "social status"],
            ["独身", "どくしん", "single, unmarried"],
            ["刺身", "さしみ", "sashimi (sliced raw fish)"]
        ]],
        ["酒", "シュ / さけ / さか / sake, alcohol, cooking sake, liquor store", [
            ["日本酒", "にほんしゅ", "sake"],
            ["お酒", "おさけ", "alcohol, liquor"],
            ["料理酒", "りょうりしゅ", "cooking sake"],
            ["酒屋", "さかや", "a liquor store"]
        ]],
        ["吸", "キュウ / す-う / breathing, inhale", [
            ["呼吸", "こきゅう", "breathing, respiration"],
            ["吸う", "すう", "breathe / inhale"]
        ]],
        ["欲", "ヨク / ほ-しい / appetite, want, motivation", [
            ["食欲", "しょくよく", "appetite"],
            ["欲しい", "ほしい", "want"],
            ["意欲", "いよく", "a will, eagerness, motivation"]
        ]],
        ["眠", "ミン / ねむ-い / ねむ-る / sleep, sleepy", [
            ["睡眠", "すいみん", "sleep"],
            ["眠い", "ねむい", "sleepy"],
            ["眠る", "ねむる", "sleep"]
        ]],
        ["疲", "つか-れる / get tired, exhaust", [
            ["疲れる", "つかれる", "get tired, exhaust"]
        ]],
        ["息", "いき / breath, son, short of breath", [
            ["息", "いき", "a breath"],
            ["息子", "むすこ", "a son"],
            ["息切れ", "いきぎれ", "be short of breath"]
        ]],
        ["呼", "コ / よ-ぶ / breathing, call", [
            ["呼吸", "こきゅう", "breathing, respiration"],
            ["呼ぶ", "よぶ", "call"]
        ]],

        # 第5週 7日目：クイズ⑤ 読みはどちら？
        ["厚", "あつ-い / thick", [
            ["厚い", "あつい", "thick"]
        ]],
        ["泣", "な-く / cry, weep", [
            ["泣く", "なく", "cry / weep"]
        ]],
        ["鳴", "な-く / な-る / chirp, ring, chime", [
            ["鳴く", "なく", "chirp / croak / bleat (etc.)"],
            ["鳴る", "なる", "ring, chime"]
        ]],
        ["初", "ショ / はじ-め / はじ-めて / first, beginning, level", [
            ["最初", "さいしょ", "first / beginning"],
            ["初め", "はじめ", "the beginning"],
            ["初級", "しょきゅう", "beginning level"],
            ["初めて", "はじめて", "first time"]
        ]],
        ["泊", "ハク / と-まる / と-める / lodging, stay, accommodate", [
            ["宿泊", "しゅくはく", "lodging"],
            ["泊まる", "とまる", "stay"],
            ["泊める", "とめる", "accommodate"]
        ]],
        ["葉", "ヨウ / は / leaf, word, language", [
            ["紅葉", "こうよう", "autumn leaves"],
            ["葉", "は", "a leaf"],
            ["言葉", "ことば", "word / language"]
        ]]
    ]
}

def generate_soma_document(subset_dict, filename="Kanji_Practice_Chapter4_5cols.docx"):
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
    generate_soma_document(CHAPTER_4_DATA, "Kanji_Practice_Chapter4_5cols.docx")
    generate_soma_document(CHAPTER_5_DATA, "Kanji_Practice_Chapter5_5cols.docx")


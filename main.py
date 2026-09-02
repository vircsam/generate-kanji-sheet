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

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
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

# Dataset covering Sets 6 to 16
DATA = {
    "Set 6": [
        ("副", "vice", [
            ("副詞", "ふくし", "adverb [grammar]"),
            ("副作用", "ふくさよう", "side effect, adverse reaction"),
            ("副社長", "ふくしゃちょう", "executive vice-president"),
            ("副産物", "ふくさんぶつ", "by-product, spin-off"),
            ("副賞", "ふくしょう", "supplementary prize, bonus prize")
        ]),
        ("算", "calculate", [
            ("予算", "よさん", "estimate (of costs), budget"),
            ("計算", "けいさん", "calculation, computation"),
            ("掛け算", "かけざん", "multiplication [mathematics]"),
            ("割り算", "わりざん", "division [mathematics]")
        ]),
        ("育", "bring up", [
            ("教育", "きょういく", "education, schooling, training"),
            ("育てる", "そだてる", "to raise, to bring up, to nurture"),
            ("育つ", "そだつ", "to be raised, to grow up"),
            ("体育", "たいいく", "physical education, PE, gym")
        ]),
        ("席", "seat", [
            ("席", "せき", "seat, location, place"),
            ("出席", "しゅっせき", "attendance, presence"),
            ("座席", "ざせき", "seat, assigned seating"),
            ("欠席", "けっせき", "absence, non-attendance"),
            ("客席", "きゃくせき", "guest seating, passenger seat")
        ]),
        ("残", "remain", [
            ("残る", "のこる", "to remain, to be left"),
            ("残念", "ざんねん", "regrettable, unfortunate, disappointing"),
            ("残す", "のこす", "to leave behind, to save, to reserve"),
            ("残り", "のこり", "remainder, remnant, the rest")
        ]),
        ("想", "idea", [
            ("想像", "そうぞう", "imagination, guess"),
            ("思想", "しそう", "thought, idea, ideology"),
            ("理想", "りそう", "ideal, ideals"),
            ("連想", "れんそう", "association of ideas"),
            ("感想", "かんそう", "impression, thoughts, comments")
        ]),
        ("線", "line", [
            ("線", "せん", "line, stripe, track, wire"),
            ("地平線", "ちへいせん", "horizon (land)"),
            ("線路", "せんろ", "railway track, railroad"),
            ("新幹線", "しんかんせん", "Shinkansen, bullet train")
        ]),
        ("農", "agriculture", [
            ("農業", "のうぎょう", "agriculture, farming"),
            ("農民", "のうみん", "farmer, peasant"),
            ("農家", "のうか", "farmer, farming family, farmhouse"),
            ("農産物", "のうさんぶつ", "agricultural produce"),
            ("農村", "のうそん", "agricultural community, farm village")
        ]),
        ("州", "state", [
            ("州", "しゅう", "state, province, county"),
            ("欧州", "おうしゅう", "Europe"),
            ("九州", "きゅうしゅう", "Kyushu (southern main island)")
        ]),
        ("念", "wish", [
            ("残念", "ざんねん", "regrettable, disappointing"),
            ("記念", "きねん", "commemoration, remembrance, memento"),
            ("観念", "かんねん", "idea, notion, concept, resignation"),
            ("概念", "がいねん", "concept, notion, conception")
        ]),
        ("象", "elephant", [
            ("対象", "たいしょう", "target, object, subject"),
            ("象", "ぞう", "elephant"),
            ("現象", "げんしょう", "phenomenon"),
            ("印象", "いんしょう", "impression"),
            ("象徴", "しょうちょう", "symbol, emblem")
        ]),
        ("助", "help", [
            ("助ける", "たすける", "to save, to rescue, to help"),
            ("援助", "えんじょ", "assistance, aid, support"),
            ("救助", "きゅうじょ", "relief, aid, rescue"),
            ("助手", "じょしゅ", "assistant, helper")
        ]),
        ("労", "labor", [
            ("労働", "ろうどう", "labor, work"),
            ("苦労", "くろう", "trouble, hardship, toil, anxiety"),
            ("ご苦労様", "ごくろうさま", "thank you for your hard work"),
            ("疲労", "ひろう", "fatigue, weariness, exhaustion")
        ]),
        ("例", "example", [
            ("例えば", "たとえば", "for example, for instance"),
            ("例", "れい", "example, instance, custom"),
            ("例外", "れいがい", "exception"),
            ("実例", "じつれい", "example, illustration"),
            ("比例", "ひれい", "proportion, proportional representation")
        ]),
        ("然", "sort of thing", [
            ("突然", "とつぜん", "abrupt, sudden, unexpected"),
            ("自然", "しぜん", "nature, natural, spontaneous"),
            ("偶然", "ぐうぜん", "coincidence, chance, accident"),
            ("全然", "ぜんぜん", "(not) at all, completely")
        ]),
        ("追", "chase", [
            ("追う", "おう", "to chase, to run after, to pursue"),
            ("追いつく", "おいつく", "to catch up with, to draw level"),
            ("追加", "ついか", "addition, supplement"),
            ("追跡", "ついせき", "pursuit, tracking, tracing")
        ]),
        ("商", "make a deal", [
            ("商品", "しょうひん", "commodity, goods, merchandise"),
            ("商売", "しょうばい", "trade, business, commerce"),
            ("商人", "しょうにん", "merchant, trader, shopkeeper"),
            ("商業", "しょうぎょう", "commerce, trade, business"),
            ("商店", "しょうてん", "shop, store")
        ]),
        ("葉", "leaf", [
            ("言葉", "ことば", "language, dialect, word, phrase"),
            ("葉", "は", "leaf, blade of grass, foliage"),
            ("紅葉", "こうよう", "autumn colours, red leaves"),
            ("落ち葉", "おちば", "fallen leaves, leaf litter")
        ]),
        ("伝", "transmit", [
            ("伝える", "つたえる", "to convey, report, transmit"),
            ("手伝う", "てつだう", "to help, assist, aid"),
            ("伝統", "でんとう", "tradition, convention"),
            ("宣伝", "せんでん", "publicity, advertisement"),
            ("伝わる", "つたわる", "to be handed down, be transmitted")
        ]),
        ("形", "shape", [
            ("人形", "にんぎょう", "doll, puppet, marionette"),
            ("形式", "けいしき", "form, format, style, formality"),
            ("形容詞", "けいようし", "adjective, i-adjective"),
            ("長方形", "ちょうほうけい", "rectangle, oblong"),
            ("正方形", "せいほうけい", "square")
        ])
    ],
    "Set 7": [
        ("景", "scenery", [
            ("光景", "こうけい", "scene, spectacle, sight, view"),
            ("風景", "ふうけい", "scenery, landscape, view"),
            ("景気", "けいき", "business conditions, economic climate"),
            ("背景", "はいけい", "background, scenery, backdrop")
        ]),
        ("落", "fall", [
            ("落す", "おとす", "to drop"),
            ("落第", "らくだい", "failure (exam), failing to advance"),
            ("墜落", "ついらく", "fall, crash (of aircraft)"),
            ("落下", "らっか", "fall, drop, descent"),
            ("落書き", "らくがき", "scrawl, scribble, graffiti")
        ]),
        ("賞", "prize", [
            ("賞", "しょう", "prize, award"),
            ("鑑賞", "かんしょう", "appreciation (of art, music)"),
            ("賞金", "しょうきん", "prize money, monetary award"),
            ("賞品", "しょうひん", "prize, trophy"),
            ("受賞", "じゅしょう", "winning a prize, receiving an award")
        ]),
        ("辺", "environs", [
            ("辺", "へん", "area, vicinity, region; side"),
            ("辺り", "あたり", "neighborhood, vicinity, nearby"),
            ("周辺", "しゅうへん", "circumference, outskirts; peripheral"),
            ("海辺", "うみべ", "beach, seashore, seaside, coast")
        ]),
        ("負", "defeat", [
            ("負ける", "まける", "to lose, to be defeated, yield"),
            ("背負う", "せおう", "to carry on back, be burdened with"),
            ("勝負", "しょうぶ", "victory or defeat, match, game")
        ]),
        ("失", "lose", [
            ("失敗", "しっぱい", "failure, mistake, blunder"),
            ("失う", "うしなう", "to lose, to miss (a chance)"),
            ("失業", "しつぎょう", "unemployment, losing job"),
            ("失望", "しつぼう", "disappointment, despair")
        ]),
        ("差", "distinction", [
            ("差別", "さべつ", "distinction, discrimination"),
            ("差", "さ", "difference, variation"),
            ("人差し指", "ひとさしゆび", "index finger, forefinger"),
            ("差し支え", "さしつかえ", "hindrance, impediment")
        ]),
        ("課", "chapter", [
            ("課長", "かちょう", "section manager, section chief"),
            ("科目", "かもく", "(school) subject, curriculum"),
            ("課", "か", "lesson, section, division"),
            ("課税", "かぜい", "taxation"),
            ("日課", "にっか", "daily routine, daily work")
        ]),
        ("末", "end", [
            ("粗末", "そまつ", "crude, plain, shabby, careless"),
            ("末", "すえ", "end, tip, top, youngest child"),
            ("末っ子", "すえっこ", "youngest child")
        ]),
        ("守", "guard", [
            ("留守", "るす", "absence, being away from home"),
            ("守る", "まもる", "to protect, guard, defend, keep"),
            ("守衛", "しゅえい", "security guard, caretaker"),
            ("守備", "しゅび", "defense (sports), guarding")
        ]),
        ("極", "extreme", [
            ("積極的", "せっきょくてき", "positive, assertive, proactive"),
            ("北極", "ほっきょく", "North Pole, the Arctic"),
            ("消極的", "しょうきょくてき", "negative, passive, half-hearted"),
            ("南極", "なんきょく", "South Pole, the Antarctic"),
            ("極端", "きょくたん", "extreme, extremity")
        ]),
        ("種", "species", [
            ("種", "たね", "seed, ingredient, secret"),
            ("種類", "しゅるい", "variety, kind, type, category"),
            ("人種", "じんしゅ", "race (of people), type of person"),
            ("一種", "いっしゅ", "a kind, a sort, a species"),
            ("各種", "かくしゅ", "every kind, all sorts")
        ]),
        ("美", "beauty", [
            ("美しい", "うつくしい", "beautiful, pretty, lovely"),
            ("美術館", "びじゅつかん", "art museum, art gallery"),
            ("美人", "びじん", "beautiful woman, belle"),
            ("美", "び", "beauty"),
            ("優美", "ゆうび", "graceful, elegant")
        ]),
        ("命", "fate", [
            ("一生懸命", "いっしょうけんめい", "very hard, with utmost effort"),
            ("生命", "せいめい", "life, existence"),
            ("命", "いのち", "life, life force, lifespan"),
            ("命令", "めいれい", "command, order, instruction")
        ]),
        ("福", "blessing", [
            ("幸福", "こうふく", "happiness, well-being, joy"),
            ("福祉", "ふくし", "welfare, social security"),
            ("福", "ふく", "good fortune, blessing, good luck"),
            ("裕福", "ゆうふく", "wealthy, rich, affluent"),
            ("祝福", "しゅくふく", "blessing, benediction")
        ]),
        ("量", "quantity", [
            ("量", "りょう", "quantity, amount, volume"),
            ("計る", "はかる", "to measure, weigh, time"),
            ("重量", "じゅうりょう", "weight"),
            ("測量", "そくりょう", "measurement, surveying"),
            ("裁量", "さいりょう", "discretion")
        ]),
        ("望", "ambition", [
            ("希望", "きぼう", "hope, wish, aspiration"),
            ("望む", "のぞむ", "to desire, wish for, see"),
            ("失望", "しつぼう", "disappointment, despair"),
            ("望み", "のぞみ", "wish, desire, prospect")
        ]),
        ("観", "outlook", [
            ("観察", "かんさつ", "observation, survey, watching"),
            ("観客", "かんきゃく", "spectator, audience"),
            ("観光", "かんこう", "sightseeing, tourism"),
            ("観測", "かんそく", "observation, survey, measurement"),
            ("観念", "かんねん", "idea, concept, notion")
        ]),
        ("察", "guess", [
            ("警察", "けいさつ", "police, police station"),
            ("観察", "かんさつ", "observation, survey"),
            ("診察", "しんさつ", "medical examination, checkup"),
            ("視察", "しさつ", "inspection, observation"),
            ("察する", "さっする", "to guess, sense, presume")
        ]),
        ("整", "organize", [
            ("整理", "せいり", "sorting, organization, putting in order"),
            ("調整", "ちょうせい", "adjustment, coordination"),
            ("整う", "ととのう", "to be prepared, be in order"),
            ("整備", "せいび", "maintenance, servicing")
        ])
    ],
    "Set 8": [
        ("横", "sideways", [
            ("横", "よこ", "horizontal, side-to-side, beside"),
            ("横断", "おうだん", "crossing, traversing"),
            ("横切る", "よこぎる", "to cross, to traverse")
        ]),
        ("型", "mould", [
            ("典型", "てんけい", "type, pattern, model, archetype"),
            ("模型", "もけい", "model, dummy, maquette"),
            ("血液型", "けつえきがた", "blood type, blood group"),
            ("体型", "たいけい", "figure, body shape, build"),
            ("原型", "げんけい", "model, prototype, archetype")
        ]),
        ("深", "deep", [
            ("深い", "ふかい", "deep, profound, dense, close"),
            ("深刻", "しんこく", "serious, severe, grave"),
            ("深夜", "しんや", "late at night"),
            ("深まる", "ふかまる", "to deepen, intensify"),
            ("深める", "ふかめる", "to deepen, heighten (transitive)")
        ]),
        ("申", "have the honor to", [
            ("申す", "もうす", "to say, to be called (humble)"),
            ("申し込む", "もうしこむ", "to apply for, propose, offer"),
            ("申請", "しんせい", "application, request, petition"),
            ("申告", "しんこく", "report, tax return, statement")
        ]),
        ("様", "way", [
            ("様々", "さまざま", "various, varied, diverse"),
            ("同様", "どうよう", "same, similar, just like"),
            ("様子", "ようす", "state of affairs, situation, appearance"),
            ("模様", "もよう", "pattern, design, condition")
        ]),
        ("港", "harbor", [
            ("空港", "くうこう", "airport"),
            ("港", "みなと", "harbour, port"),
            ("港湾", "こうわん", "harbour, port"),
            ("寄港", "きこう", "calling at a port, stopover"),
            ("漁港", "ぎょこう", "fishing port")
        ]),
        ("達", "accomplished", [
            ("友達", "ともだち", "friend, companion"),
            ("発達", "はったつ", "development, growth, progress"),
            ("達する", "たっする", "to reach, get to, arrive at"),
            ("配達", "はいたつ", "delivery"),
            ("上達", "じょうたつ", "improvement, progress (in skill)")
        ]),
        ("良", "good", [
            ("仲良し", "なかよし", "close friendship, good friend"),
            ("良心", "りょうしん", "conscience"),
            ("改良", "かいりょう", "improvement, reform"),
            ("善良", "ぜんりょう", "good-natured, virtuous, honest")
        ]),
        ("谷", "valley", [
            ("谷", "たに", "valley, ravine, gorge"),
            ("谷間", "たにま", "valley, gorge, chasm, gap"),
            ("峡谷", "きょうこく", "gorge, ravine, canyon"),
            ("渓谷", "けいこく", "valley (with river), gorge")
        ]),
        ("候", "climate", [
            ("天候", "てんこう", "weather"),
            ("気候", "きこう", "climate"),
            ("候補", "こうほ", "candidate, choice"),
            ("立候補", "りっこうほ", "standing as a candidate"),
            ("居候", "いそうろう", "lodger, freeloader, sponger")
        ]),
        ("史", "history", [
            ("歴史", "れきし", "history"),
            ("女史", "じょし", "lady (high status), Ms."),
            ("史料", "しりょう", "historical materials, archives"),
            ("日本史", "にほんし", "Japanese history"),
            ("史跡", "しせき", "historic landmark, historic site")
        ]),
        ("階", "storey", [
            ("階段", "かいだん", "stairs, stairway"),
            ("二階建て", "にかいだて", "two-storied building"),
            ("段階", "だんかい", "stage, step, phase, level"),
            ("階級", "かいきゅう", "class, rank, grade"),
            ("階層", "かいそう", "hierarchy, stratum, class")
        ]),
        ("満", "full", [
            ("満足", "まんぞく", "satisfaction, contentment"),
            ("満ちる", "みちる", "to fill, become full, mature"),
            ("不満", "ふまん", "dissatisfaction, discontent")
        ]),
        ("敗", "failure", [
            ("失敗", "しっぱい", "failure, mistake, blunder"),
            ("勝敗", "しょうはい", "victory or defeat, outcome"),
            ("腐敗", "ふはい", "decomposition, corruption"),
            ("敗戦", "はいせん", "defeat, lost battle"),
            ("敗北", "はいぼく", "defeat, setback"),
            ("敗れる", "やぶれる", "to be defeated, to lose")
        ]),
        ("管", "pipe", [
            ("管理", "かんり", "control, management"),
            ("管", "かん", "pipe, tube, duct"),
            ("血管", "けっかん", "blood vessel, vein"),
            ("保管", "ほかん", "custody, safekeeping, storage"),
            ("管轄", "かんかつ", "jurisdiction, control")
        ]),
        ("兵", "soldier", [
            ("兵隊", "へいたい", "military, troops, soldier"),
            ("兵士", "へいし", "soldier"),
            ("兵器", "へいき", "arms, weapon, ordnance"),
            ("核兵器", "かくへいき", "nuclear weapon"),
            ("徴兵", "ちょうへい", "conscription, military draft")
        ]),
        ("器", "utensil", [
            ("機械", "きかい", "machine, mechanism, appliance"),
            ("武器", "ぶき", "weapon, arms, asset"),
            ("器用", "きよう", "skillful, adroit, dexterous"),
            ("食器", "しょっき", "tableware"),
            ("容器", "ようき", "container, receptacle, vessel")
        ]),
        ("路", "path", [
            ("道路", "どうろ", "road, highway"),
            ("線路", "せんろ", "railway track, railroad"),
            ("通路", "つうろ", "passage, pathway, aisle"),
            ("進路", "しんろ", "route, course, future path"),
            ("回路", "かいろ", "circuit [electricity]")
        ]),
        ("科", "department", [
            ("科学", "かがく", "science"),
            ("教科書", "きょうかしょ", "textbook, coursebook"),
            ("科目", "かもく", "school subject, curriculum"),
            ("外科", "げか", "surgery department"),
            ("学科", "がっか", "academic course, department")
        ]),
        ("細", "dainty", [
            ("細い", "ほそい", "thin, slender, fine"),
            ("細かい", "こまかい", "small, fine, minute, trivial"),
            ("詳細", "しょうさい", "details, particulars"),
            ("細胞", "さいぼう", "cell [biology]")
        ])
    ],
    "Set 9": [
        ("積", "volume", [
            ("積極的", "せっきょくてき", "positive, assertive, proactive"),
            ("積もる", "つもる", "to pile up, to accumulate"),
            ("積む", "つむ", "to pile up, stack, load"),
            ("面積", "めんせき", "area, square measure, size")
        ]),
        ("丸", "round", [
            ("丸", "まる", "circle, whole, complete"),
            ("丸める", "まるめる", "to make round, roll up, curl"),
            ("日の丸", "ひのまる", "red circle of the Sun")
        ]),
        ("他", "other", [
            ("他", "ほか", "other"),
            ("他人", "たにん", "another person, others, stranger"),
            ("その他", "そのほか", "the rest, the others, besides"),
            ("他方", "たほう", "on the other hand")
        ]),
        ("録", "record", [
            ("記録", "きろく", "record, document, minutes"),
            ("録音", "ろくおん", "audio recording"),
            ("登録", "とうろく", "registration, entry, record"),
            ("目録", "もくろく", "catalogue, inventory, contents"),
            ("付録", "ふろく", "appendix, supplement")
        ]),
        ("省", "ministry", [
            ("省く", "はぶく", "to omit, exclude, economize"),
            ("省略", "しょうりゃく", "omission, abbreviation"),
            ("反省", "はんせい", "reflection, introspection, regret"),
            ("省みる", "かえりみる", "to reflect on, contemplate"),
            ("大蔵省", "おおくらしょう", "Ministry of Finance (former)")
        ]),
        ("橋", "bridge", [
            ("橋", "はし", "bridge"),
            ("鉄橋", "てっきょう", "railway bridge, iron bridge"),
            ("桟橋", "さんばし", "wharf, jetty, pier"),
            ("橋渡し", "はしわたし", "mediation, go-between"),
            ("歩道橋", "ほどうきょう", "pedestrian overpass")
        ]),
        ("岸", "beach", [
            ("海岸", "かいがん", "seashore, coast, seaside, beach"),
            ("岸", "きし", "bank, coast, shore"),
            ("沿岸", "えんがん", "coast, shore, coastal area"),
            ("対岸", "たいがん", "opposite shore"),
            ("岸壁", "がんぺき", "quay, wharf, cliff wall")
        ]),
        ("客", "guest", [
            ("客", "きゃく", "guest, visitor, customer"),
            ("乗客", "じょうきゃく", "passenger"),
            ("観客", "かんきゃく", "spectator, audience"),
            ("客席", "きゃくせき", "audience seat, spectator seat")
        ]),
        ("周", "circumference", [
            ("周り", "まわり", "circumference, surroundings"),
            ("周囲", "しゅうい", "surroundings, environs"),
            ("周辺", "しゅうへん", "outskirts, peripheral"),
            ("円周", "えんしゅう", "circumference"),
            ("周期", "しゅうき", "period, cycle")
        ]),
        ("材", "lumber", [
            ("材料", "ざいりょう", "materials, ingredients, data"),
            ("木材", "もくざい", "lumber, timber, wood"),
            ("素材", "そざい", "ingredient, raw material"),
            ("取材", "しゅざい", "collecting info, reporting"),
            ("人材", "じんざい", "human resources, talent")
        ]),
        ("登", "ascend", [
            ("登山", "とざん", "mountain climbing, ascent"),
            ("登場", "とうじょう", "entrance on stage, emergence"),
            ("登録", "とうろく", "registration, entry, record"),
            ("登校", "とうこう", "attendance at school"),
            ("山登り", "やまのぼり", "mountain climbing")
        ]),
        ("健", "healthy", [
            ("健康", "けんこう", "health, sound, fit"),
            ("保健", "ほけん", "preservation of health, hygiene"),
            ("健全", "けんぜん", "healthy, sound, wholesome"),
            ("健やか", "すこやか", "vigorous, healthy, sound"),
            ("穏健", "おんけん", "moderate, temperate, sensible")
        ]),
        ("戸", "door", [
            ("戸", "と", "door, shutter"),
            ("井戸", "いど", "water well"),
            ("雨戸", "あまど", "sliding storm shutter"),
            ("戸棚", "とだな", "cupboard, cabinet, closet"),
            ("戸籍", "こせき", "family register")
        ]),
        ("速", "quick", [
            ("速い", "はやい", "fast, quick, rapid, speedy"),
            ("急速", "きゅうそく", "rapid (progress)"),
            ("速度", "そくど", "speed, velocity, pace"),
            ("高速", "こうそく", "high-speed, express")
        ]),
        ("飛", "fly", [
            ("飛行機", "ひこうき", "airplane, aircraft"),
            ("飛ぶ", "とぶ", "to fly, to leap, to scatter"),
            ("飛ばす", "とばす", "to let fly, fire, skip over")
        ]),
        ("殺", "kill", [
            ("殺す", "ころす", "to kill, murder, suppress"),
            ("自殺", "じさつ", "suicide"),
            ("殺人", "さつじん", "murder, homicide"),
            ("暗殺", "あんさつ", "assassination"),
            ("殺虫剤", "さっちゅうざい", "insecticide")
        ]),
        ("央", "center", [
            ("中央", "ちゅうおう", "centre, middle, heart, capital"),
            ("道央", "どうおう", "central Hokkaido"),
            ("震央", "しんおう", "epicentre (of earthquake)"),
            ("月央", "げつおう", "middle of the month"),
            ("年央", "ねんおう", "mid-year")
        ]),
        ("号", "nickname", [
            ("番号", "ばんごう", "number, series of digits"),
            ("信号", "しんごう", "traffic light, signal"),
            ("記号", "きごう", "sign, symbol, mark"),
            ("符号", "ふごう", "sign, code, math symbol"),
            ("年号", "ねんごう", "era name")
        ]),
        ("単", "simple", [
            ("簡単", "かんたん", "simple, easy, brief"),
            ("単位", "たんい", "unit, denomination, credit"),
            ("単に", "たんに", "simply, merely, only"),
            ("単語", "たんご", "word, vocabulary")
        ]),
        ("竹", "bamboo", [
            ("竹", "たけ", "bamboo"),
            ("竹林", "ちくりん", "bamboo thicket, bamboo grove"),
            ("竹刀", "しない", "bamboo sword (kendo)"),
            ("爆竹", "ばくちく", "firecracker")
        ])
    ],
    "Set 10": [
        ("完", "perfect", [
            ("完全", "かんぜん", "perfect, complete"),
            ("完成", "かんせい", "completion, perfection"),
            ("完了", "かんりょう", "completion, perfect tense"),
            ("完璧", "かんぺき", "perfect, flawless"),
            ("完結", "かんけつ", "conclusion, completion")
        ]),
        ("競", "emulate", [
            ("競争", "きょうそう", "competition, rivalry, race"),
            ("競技", "きょうぎ", "game, match, sporting event"),
            ("競馬", "けいば", "horse racing"),
            ("競う", "きそう", "to compete, vie, contend"),
            ("競る", "せる", "to compete, to bid")
        ]),
        ("給", "salary", [
            ("供給", "きょうきゅう", "supply, provision"),
            ("給料", "きゅうりょう", "salary, wages, pay"),
            ("支給", "しきゅう", "provision, allowance, grant"),
            ("給与", "きゅうよ", "pay, wages, salary"),
            ("月給", "げっきゅう", "monthly salary")
        ]),
        ("根", "root", [
            ("根", "ね", "root, source, origin, cause"),
            ("根拠", "こんきょ", "basis, grounds, authority"),
            ("根気", "こんき", "patience, perseverance"),
            ("根底", "こんてい", "root, basis, foundation"),
            ("根回し", "ねまわし", "laying groundwork, consensus")
        ]),
        ("苦", "suffering", [
            ("苦い", "にがい", "bitter"),
            ("苦労", "くろう", "trouble, hardship, difficulty"),
            ("苦痛", "くつう", "pain, agony, distress"),
            ("苦手", "にがて", "poor at, not one's cup of tea"),
            ("苦しい", "くるしい", "painful, difficult")
        ]),
        ("園", "park", [
            ("公園", "こうえん", "public park"),
            ("動物園", "どうぶつえん", "zoo, zoological gardens"),
            ("園芸", "えんげい", "horticulture, gardening"),
            ("幼稚園", "ようちえん", "kindergarten, preschool"),
            ("遊園地", "ゆうえんち", "amusement park"),
            ("園", "その", "garden, park")
        ]),
        ("具", "tool", [
            ("道具", "どうぐ", "tool, implement, instrument"),
            ("具合", "ぐあい", "condition, state, health"),
            ("家具", "かぐ", "furniture"),
            ("具体", "ぐたい", "concreteness, embodiment")
        ]),
        ("歴", "curriculum", [
            ("歴史", "れきし", "history"),
            ("経歴", "けいれき", "personal history, career"),
            ("学歴", "がくれき", "academic background"),
            ("遍歴", "へんれき", "travels, pilgrimage"),
            ("履歴書", "りれきしょ", "resume, CV")
        ]),
        ("辞", "resign", [
            ("辞書", "じしょ", "dictionary, lexicon"),
            ("辞典", "じてん", "dictionary, lexicon"),
            ("辞める", "やめる", "to resign, quit, retire"),
            ("お辞儀", "おじぎ", "bow, bowing"),
            ("百科事典", "ひゃっかじてん", "encyclopedia")
        ]),
        ("馬", "horse", [
            ("馬", "うま", "horse"),
            ("競馬", "けいば", "horse racing"),
            ("馬車", "ばしゃ", "coach, carriage, wagon"),
            ("馬力", "ばりき", "horsepower, energy, vitality")
        ]),
        ("愛", "love", [
            ("愛", "あい", "love, affection, care"),
            ("愛情", "あいじょう", "love, affection"),
            ("愛する", "あいする", "to love"),
            ("可愛らしい", "かわいらしい", "lovely, sweet, pretty, cute"),
            ("恋愛", "れんあい", "love, romance, passion")
        ]),
        ("未", "un-", [
            ("未来", "みらい", "(distant) future"),
            ("未満", "みまん", "less than, under, below"),
            ("未熟", "みじゅく", "unripe, immature, unskilled"),
            ("未婚", "みこん", "unmarried, single"),
            ("未払い", "みはらい", "unpaid, overdue")
        ]),
        ("航", "navigate", [
            ("航空", "こうくう", "aviation, flying"),
            ("航海", "こうかい", "voyage, navigation, sailing"),
            ("運航", "うんこう", "operation (of ship/aircraft)"),
            ("航路", "こうろ", "route, flight course"),
            ("就航", "しゅうこう", "entering service (plane/ship)")
        ]),
        ("冷", "cool", [
            ("冷たい", "つめたい", "cold (to touch), coldhearted"),
            ("冷蔵庫", "れいぞうこ", "refrigerator, fridge"),
            ("冷える", "ひえる", "to grow cold, get chilly"),
            ("冷房", "れいぼう", "air conditioning, cooling"),
            ("冷静", "れいせい", "calm, composure, coolness")
        ]),
        ("鉄", "iron", [
            ("地下鉄", "ちかてつ", "subway, metro"),
            ("鉄道", "てつどう", "railroad, railway"),
            ("鉄", "てつ", "iron (Fe), steel"),
            ("鉄砲", "てっぽう", "gun, firearm")
        ]),
        ("類", "sort", [
            ("種類", "しゅるい", "variety, kind, type, category"),
            ("人類", "じんるい", "mankind, humanity"),
            ("書類", "しょるい", "document, papers"),
            ("分類", "ぶんるい", "classification, sorting"),
            ("親類", "しんるい", "relative, relation, kin")
        ]),
        ("児", "newborn babe", [
            ("幼児", "ようじ", "young child, toddler"),
            ("児童", "じどう", "children, juvenile"),
            ("育児", "いくじ", "childcare, upbringing"),
            ("小児科", "しょうにか", "pediatrics"),
            ("双生児", "そうせいじ", "twins"),
            ("園児", "えんじ", "kindergartener")
        ]),
        ("印", "stamp", [
            ("印", "しるし", "mark, sign, symbol, token"),
            ("印刷", "いんさつ", "printing"),
            ("印象", "いんしょう", "impression"),
            ("目印", "めじるし", "mark, landmark, guide")
        ]),
        ("王", "king", [
            ("王", "おう", "king, ruler, tycoon"),
            ("女王", "じょおう", "queen, female champion"),
            ("王子", "おうじ", "prince"),
            ("王様", "おうさま", "king (respectful)"),
            ("王女", "おうじょ", "princess")
        ]),
        ("返", "return", [
            ("返す", "かえす", "to return, pay back, do again"),
            ("返事", "へんじ", "reply, answer, response"),
            ("繰り返す", "くりかえす", "to repeat, do over"),
            ("裏返す", "うらがえす", "to turn inside out, flip over"),
            ("返る", "かえる", "to return, come back")
        ])
    ]
}

def generate_kanji_document(filename="Kanji_Practice_Sheets_Sets_6_to_10.docx"):
    doc = Document()
    
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(0.4)
        s.bottom_margin = Inches(0.4)
        s.left_margin = Inches(0.5)
        s.right_margin = Inches(0.5)

    COL_WIDTHS = [Inches(1.8)] + [Inches(0.8)] * 7
    
    for set_name, kanji_list in DATA.items():
        for kanji, meaning, vocab_list in kanji_list:
            chunk_size = 3
            chunks = [vocab_list[i:i+chunk_size] for i in range(0, len(vocab_list), chunk_size)]
            
            for chunk_idx, chunk in enumerate(chunks):
                p_set = doc.add_paragraph()
                p_set.paragraph_format.space_before = Pt(2)
                p_set.paragraph_format.space_after = Pt(2)
                run_set = p_set.add_run(f"{set_name}" if chunk_idx == 0 else f"{set_name} (cont.)")
                run_set.font.name = "Georgia"
                run_set.font.size = Pt(14)
                run_set.font.bold = True
                run_set.font.color.rgb = RGBColor(0x8C, 0x2A, 0x1E)

                header_tbl = doc.add_table(rows=1, cols=1)
                header_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                h_cell = header_tbl.rows[0].cells[0]
                h_cell.width = Inches(7.4)
                set_cell_background(h_cell, "F3EFE9")
                set_cell_margins(h_cell, top=80, bottom=80, left=150, right=150)
                
                tcPr = h_cell._element.get_or_add_tcPr()
                borders = parse_xml(
                    f'<w:tcBorders {nsdecls("w")}>'
                    f'<w:left w:val="single" w:sz="36" w:space="0" w:color="8C2A1E"/>'
                    f'<w:top w:val="none"/>'
                    f'<w:right w:val="none"/>'
                    f'<w:bottom w:val="none"/>'
                    f'</w:tcBorders>'
                )
                tcPr.append(borders)

                hp = h_cell.paragraphs[0]
                hp.paragraph_format.space_before = Pt(2)
                hp.paragraph_format.space_after = Pt(1)
                k_run = hp.add_run(kanji)
                k_run.font.name = "MS Mincho"
                k_run.font.size = Pt(26)
                k_run.font.bold = True

                m_p = h_cell.add_paragraph()
                m_p.paragraph_format.space_before = Pt(0)
                m_p.paragraph_format.space_after = Pt(2)
                m_run = m_p.add_run(meaning)
                m_run.font.name = "Georgia"
                m_run.font.size = Pt(10)
                m_run.font.italic = True
                m_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

                doc.add_paragraph().paragraph_format.space_after = Pt(4)

                table = doc.add_table(rows=4, cols=8)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                set_table_borders(table, color="D3C7B6")

                hdr_row = table.rows[0]
                hdr_titles = ["WORD", "1", "2", "3", "4", "5", "6", "7"]
                for i, cell in enumerate(hdr_row.cells):
                    cell.width = COL_WIDTHS[i]
                    set_cell_background(cell, "EDE5D8")
                    set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    r = p.add_run(hdr_titles[i])
                    r.font.name = "Arial"
                    r.font.size = Pt(8.5)
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(0x77, 0x77, 0x77)

                for r_idx in range(3):
                    row = table.rows[r_idx + 1]
                    trPr = row._element.get_or_add_trPr()
                    trHeight = parse_xml(f'<w:trHeight {nsdecls("w")} w:val="1800" w:hRule="atLeast"/>')
                    trPr.append(trHeight)

                    info_cell = row.cells[0]
                    info_cell.width = COL_WIDTHS[0]
                    set_cell_margins(info_cell, top=120, bottom=120, left=120, right=80)
                    info_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                    if r_idx < len(chunk):
                        w_kanji, w_hira, w_eng = chunk[r_idx]
                        ip = info_cell.paragraphs[0]
                        ip.paragraph_format.space_before = Pt(0)
                        ip.paragraph_format.space_after = Pt(2)
                        
                        w_run = ip.add_run(w_kanji)
                        w_run.font.name = "MS Mincho"
                        w_run.font.size = Pt(20)
                        w_run.font.bold = True
                        
                        hp = info_cell.add_paragraph()
                        hp.paragraph_format.space_before = Pt(0)
                        hp.paragraph_format.space_after = Pt(2)
                        h_run = hp.add_run(w_hira)
                        h_run.font.name = "Hiragino Mincho ProN"
                        h_run.font.size = Pt(9.5)
                        h_run.font.color.rgb = RGBColor(0x8C, 0x6B, 0x3E)

                        ep = info_cell.add_paragraph()
                        ep.paragraph_format.space_before = Pt(0)
                        ep.paragraph_format.space_after = Pt(0)
                        e_run = ep.add_run(w_eng)
                        e_run.font.name = "Arial"
                        e_run.font.size = Pt(7.5)
                        e_run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
                    else:
                        info_cell.paragraphs[0].text = ""

                    for c_idx in range(1, 8):
                        p_cell = row.cells[c_idx]
                        p_cell.width = COL_WIDTHS[c_idx]
                        p_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                doc.add_page_break()

    doc.save(filename)
    print(f"Successfully generated: {filename}")

if __name__ == "__main__":
    generate_kanji_document()

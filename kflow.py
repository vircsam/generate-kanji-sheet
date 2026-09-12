import os
import re
import math
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# 1. Register Font
FONT_PATH = '/System/Library/Fonts/Supplemental/Arial Unicode.ttf'
pdfmetrics.registerFont(TTFont('ArialUnicode', FONT_PATH))

# 2. Raw Data Parsing
RAW_DATA = """
┏━ 水  (water 氵)  [25 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 満  [man]  full, satisfied
┣━ 港  [kō / minato]  harbor, port
┣━ 準  [jun]  prepare, standard
┣━ 温  [on / atata(kai)]  warm, temperature
┣━ 法  [hō]  law, method
┣━ 減  [gen / he(ru), he(rasu)]  decrease
┣━ 湯  [yu]  hot water
┣━ 氷  [kōri]  ice
┣━ 混  [kon / ma(zeru)]  mix
┣━ 渡  [wata(ru), wata(su)]  cross, hand over
┣━ 濃  [ko(i)]  thick, dark
┣━ 消  [shō / ki(eru), ke(su)]  extinguish, disappear
┣━ 汗  [ase]  sweat
┣━ 涙  [namida]  tear
┣━ 治  [ji/chi / nao(ru), nao(su)]  cure, heal
┣━ 汚  [kitanai, yogo(reru)]  dirty
┣━ 泣  [na(ku)]  cry
┣━ 泊  [haku / to(maru), to(meru)]  stay overnight
┣━ 波  [ha / nami]  wave
┣━ 決  [ketsu / ki(meru), ki(maru)]  decide
┣━ 済  [sai / su(mu)]  finish, settle
┣━ 活  [katsu]  life, activity
┣━ 流  [naga(reru), naga(su)]  flow
┣━ 油  [yu / abura]  oil
┗━ 湖  [ko / mizuumi]  lake

┏━ 人  (person 亻)  [17 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 件  [ken]  matter, case
┣━ 保  [ho]  preserve, protect
┣━ 他  [ta / hoka]  other
┣━ 倍  [bai]  times, double
┣━ 優  [yū / yasa(shii)]  gentle, superior
┣━ 係  [kei / kakari]  person in charge, relation
┣━ 伺  [ukaga(u)]  ask, visit
┣━ 備  [bi / sona(eru)]  prepare, provide
┣━ 値  [ne]  value, price
┣━ 例  [rei / tato(eru)]  example
┣━ 個  [ko]  individual, counter
┣━ 停  [tei]  stop
┣━ 伝  [den / tsuta(eru)]  convey, transmit
┣━ 価  [ka]  value, price
┣━ 付  [tsu(keru), tsu(ku)]  attach, add
┣━ 側  [soku / gawa]  side
┗━ 信  [shin]  trust, believe

┏━ 糸  (thread 糸)  [13 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 線  [sen]  line
┣━ 約  [yaku]  promise, approximately
┣━ 絵  [kai / e]  picture, painting
┣━ 組  [kumi, ku(mu)]  group, assemble
┣━ 続  [zoku / tsuzu(ku), tsuzu(keru)]  continue
┣━ 級  [kyū]  grade, class
┣━ 紅  [kō / beni]  crimson, red
┣━ 緑  [ryoku / midori]  green
┣━ 結  [ketsu / musu(bu)]  tie, connect
┣━ 経  [kei]  pass through, manage
┣━ 練  [ren]  practice, train
┣━ 細  [hoso(i), koma(kai)]  thin, detailed
┗━ 絡  [raku]  entangle, contact

┏━ 木  (tree 木)  [13 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 横  [ō / yoko]  horizontal, side
┣━ 案  [an]  plan, proposal
┣━ 様  [yō / sama]  manner, appearance
┣━ 格  [kaku]  status, standard
┣━ 枚  [mai]  counter for flat objects
┣━ 機  [ki]  machine, opportunity
┣━ 査  [sa]  investigate
┣━ 果  [ka]  result, fruit
┣━ 材  [zai]  material
┣━ 束  [soku / taba]  bundle
┣━ 未  [mi]  not yet, future
┣━ 末  [matsu]  end
┗━ 橋  [kyō / hashi]  bridge

┏━ 辶  (movement 辶)  [12 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 速  [soku / haya(i)]  fast, speed
┣━ 過  [ka / su(giru)]  pass, exceed
┣━ 遊  [yū / aso(bu)]  play
┣━ 返  [hen / kae(su)]  return
┣━ 造  [zō / tsuku(ru)]  make, manufacture
┣━ 込  [ko(mu)]  crowded, include
┣━ 達  [tatsu]  reach, attain
┣━ 連  [ren / tsu(reru)]  connect, take along
┣━ 違  [chiga(u), chiga(eru)]  differ, wrong
┣━ 適  [teki]  suitable, appropriate
┣━ 選  [sen / era(bu)]  choose
┗━ 遅  [chi / oso(i), oku(reru)]  late, slow

┏━ 口  (mouth 口)  [11 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 向  [kō / mu(ko), mu(ka), mu(ki)]  direction, face toward
┣━ 号  [gō]  number, signal
┣━ 各  [kaku]  each, every
┣━ 召  [shō / me(su)]  call, summon
┣━ 商  [shō]  commerce
┣━ 告  [koku]  tell, announce
┣━ 君  [kun / kimi]  you
┣━ 吸  [kyū / su(u)]  inhale, breathe
┣━ 吹  [fu(ku)]  blow
┣━ 否  [hi]  deny, negative
┗━ 呼  [ko / yo(bu)]  call

┏━ 手  (hand 扌)  [10 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 押  [ō / o(su), o(saeru)]  push, press
┣━ 捨  [su(teru)]  throw away
┣━ 指  [shi / yubi]  finger, point
┣━ 折  [setsu / o(ru), o(reru)]  fold, break
┣━ 授  [ju]  teach, give
┣━ 払  [hara(u)]  pay
┣━ 接  [setsu]  connect, touch
┣━ 換  [kan / ka(eru)]  exchange
┣━ 投  [tō / na(geru)]  throw
┗━ 技  [gi]  skill, technique

┏━ 宀  (roof 宀)  [10 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 定  [tei]  determine, fixed
┣━ 客  [kyaku]  customer, guest
┣━ 察  [satsu]  observe, infer
┣━ 完  [kan]  complete
┣━ 寝  [ne(ru)]  sleep
┣━ 宿  [shuku / yado]  lodging, inn
┣━ 容  [yō]  contents, contain
┣━ 実  [jitsu]  truth, reality
┣━ 守  [shu/su / mamo(ru)]  protect
┗━ 宅  [taku]  home, house

┏━ 心  (heart 心/忄)  [8 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 快  [kai]  pleasant, fast
┣━ 念  [nen]  thought, concern
┣━ 想  [sō]  thought, idea
┣━ 感  [kan / kan(jiru)]  feel, emotion
┣━ 必  [hitsu / kanara(zu)]  certainly, necessarily
┣━ 性  [sei]  nature, gender
┣━ 情  [jō]  feeling, circumstance
┗━ 息  [soku / iki]  breath

┏━ 日  (sun 日)  [7 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 普  [fu]  ordinary, general
┣━ 晩  [ban]  evening, night
┣━ 昔  [mukashi]  past, long ago
┣━ 晴  [ha(reru)]  clear, sunny
┣━ 暖  [dan / atata(kai)]  warm
┣━ 昨  [saku]  yesterday, previous
┗━ 易  [eki / yasa(shii)]  easy, change

┏━ 火  (fire 火/灬)  [6 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 無  [mu / na(i)]  none, without
┣━ 煙  [en / kemuri]  smoke
┣━ 焼  [ya(ku), ya(keru)]  burn, bake
┣━ 点  [ten]  point, dot
┣━ 熱  [netsu / atsu(i)]  heat, fever
┗━ 営  [ei]  operate, manage

┏━ 言  (speech 言)  [6 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 認  [nin / mitomeru]  recognize, approve
┣━ 課  [ka]  section, lesson
┣━ 議  [gi]  discuss, deliberation
┣━ 調  [chō / shira(beru)]  investigate, condition
┣━ 記  [ki]  record, write
┗━ 警  [kei]  warn, police

┏━ 艹  (grass 艹)  [6 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 若  [waka(i)]  young
┣━ 蔵  [zō]  store, warehouse
┣━ 薄  [usu(i)]  thin, light
┣━ 苦  [ku / kuru(shii), niga(i)]  painful, bitter
┣━ 葉  [yō / ha]  leaf
┗━ 荷  [ka]  baggage, load

┏━ 阝L  (hill (left) 阝)  [6 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 階  [kai]  floor, level
┣━ 険  [ken]  dangerous, steep
┣━ 防  [bō / fuse(gu)]  prevent, defend
┣━ 限  [gen / kagi(ru)]  limit
┣━ 降  [kō / o(riru), fu(ru)]  descend, fall
┗━ 際  [sai]  occasion, edge

┏━ 力  (power 力)  [6 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 助  [jo / tasukeru]  help, rescue
┣━ 勤  [kin / tsuto(meru)]  work, serve
┣━ 務  [mu]  duty, work
┣━ 加  [ka / kuwa(eru)]  add
┣━ 募  [bo]  recruit
┗━ 労  [rō]  labor, hardship

┏━ 攵  (action/tap 攵)  [6 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 改  [kai / arata(meru)]  reform, revise
┣━ 故  [ko]  cause, former
┣━ 政  [sei]  politics, government
┣━ 救  [kyū / suku(u)]  save, rescue
┣━ 数  [sū / kazu, kazo(eru)]  number, count
┗━ 整  [sei]  arrange, organize

┏━ 貝  (shell/money 貝)  [5 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 費  [hi]  expense, cost
┣━ 賞  [shō]  prize, reward
┣━ 販  [han]  sell
┣━ 負  [fu / ma(keru)]  lose
┗━ 貿  [bō]  trade

┏━ 巾  (cloth 巾)  [5 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 常  [jō]  usual, ordinary
┣━ 席  [seki]  seat
┣━ 師  [shi]  teacher, master
┣━ 希  [ki]  hope
┗━ 帯  [tai / obi]  belt, zone

┏━ 衣  (clothes 衣)  [5 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 衣  [i]  clothing
┣━ 裏  [ura]  back, reverse
┣━ 製  [sei]  manufacture, made
┣━ 袋  [tai / fukuro]  bag
┗━ 表  [hyō / omote, arawa(su)]  surface, express

┏━ 示  (altar 示/礻)  [5 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 礼  [rei]  thanks, courtesy
┣━ 祭  [sai / matsuri]  festival
┣━ 神  [shin/jin / kami]  god, divine
┣━ 禁  [kin]  prohibit, forbid
┗━ 示  [ji / shime(su)]  show, indicate

┏━ 土  (earth 土)  [5 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 在  [zai]  exist, be present
┣━ 型  [kata]  type, model
┣━ 増  [zō / fu(eru), fu(yasu)]  increase
┣━ 塩  [en / shio]  salt
┗━ 報  [hō]  report, information

┏━ 禾  (grain 禾)  [5 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 税  [zei]  tax
┣━ 種  [shu / tane]  kind, seed
┣━ 科  [ka]  department, subject
┣━ 移  [i / utsu(ru), utsu(su)]  move, transfer
┗━ 利  [ri]  advantage, use

┏━ 頁  (head/page 頁)  [4 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 願  [nega(u)]  wish, request
┣━ 類  [rui]  kind, type
┣━ 額  [gaku]  amount, sum
┗━ 預  [yo / azu(keru)]  deposit, entrust

┏━ 月  (moon/flesh 月)  [4 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 期  [ki]  period, term
┣━ 望  [bō / nozomu]  desire, hope
┣━ 育  [iku / soda(tsu), soda(teru)]  raise, grow
┗━ 勝  [shō / ka(tsu)]  win

┏━ 田  (field 田)  [4 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 由  [yū]  reason, freedom
┣━ 番  [ban]  number, turn
┣━ 申  [shin / mō(su)]  say, apply
┗━ 留  [ryū/ru / to(meru)]  stay, retain

┏━ 囗  (enclosure 囗)  [4 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 園  [en]  garden, park
┣━ 困  [koma(ru)]  trouble
┣━ 団  [dan]  group
┗━ 因  [in]  cause

┏━ 亡  ('die/lost' 亡 family)  [3 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 亡  [bō / na(kunaru)]  die, deceased
┣━ 忙  [bō / isoga(shii)]  busy
┗━ 忘  [bō / wasu(reru)]  forget

┏━ 石  (stone 石)  [3 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 確  [kaku / tashi(ka), tashi(kameru)]  certain, confirm
┣━ 石  [seki/se(tsu) / ishi]  stone
┗━ 砂  [sa / suna]  sand

┏━ 女  (woman 女)  [3 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 婦  [fu]  woman, wife
┣━ 妻  [sai / tsuma]  wife
┗━ 婚  [kon]  marriage

┏━ 目  (eye 目)  [3 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 直  [choku / nao(ru), nao(su)]  straight, fix
┣━ 眠  [min / nemu(i), nemu(ru)]  sleep, sleepy
┗━ 相  [sō/shō / ai]  mutual, partner

┏━ 雨  (rain 雨)  [3 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 雲  [kumo]  cloud
┣━ 雪  [setsu / yuki]  snow
┗━ 震  [shin]  shake, earthquake

┏━ 欠  (lack/yawn 欠)  [3 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 次  [ji / tsugi]  next
┣━ 欠  [ketsu / ka(keru)]  lack, miss
┗━ 欲  [yoku / hoshii]  desire, want

┏━ 刀  (knife 刀/刂)  [3 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 割  [wa(ru), wa(reru)]  divide, break
┣━ 券  [ken]  ticket, coupon
┗━ 初  [sho / haji(me), haji(mete)]  first, beginning

┏━ 寸  (measure 寸)  [3 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 寺  [ji / tera]  temple
┣━ 対  [tai]  opposite, versus
┗━ 専  [sen]  specialty, exclusive

┏━ 竹  (bamboo 竹)  [3 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 箱  [hako]  box
┣━ 笑  [wara(u), e(mu)]  laugh, smile
┗━ 簡  [kan]  simple

┏━ 广  (cliff 广)  [3 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 座  [za / suwaru]  sit, seat
┣━ 府  [fu]  government office
┗━ 庫  [ko]  storehouse

┏━ 又  ('hand/again' 又)  [3 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 受  [ju / u(keru)]  receive
┣━ 取  [to(ru)]  take
┗━ 反  [han]  opposite, against

┏━ 曰  ('say' 曰)  [3 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 曲  [kyoku / ma(garu), ma(geru)]  bend, music
┣━ 最  [sai / motto(mo)]  most, utmost
┗━ 替  [ka(eru)]  replace, exchange

┏━ 大  (big 大)  [3 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 夫  [fu/fū / otto]  husband
┣━ 失  [shitsu]  lose, fail
┗━ 奥  [oku]  inner part

┏━ 馬  (horse 馬)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 馬  [ba / uma]  horse
┗━ 駐  [chū]  station, park

┏━ 阝R  (village (right) 阝)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 郵  [yū]  mail, postal
┗━ 部  [bu]  section, part

┏━ 王  (jade 王)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 現  [gen / arawa(reru)]  appear, present
┗━ 球  [kyū]  ball, sphere

┏━ 米  (rice 米)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 米  [bei / kome]  rice, America
┗━ 粉  [fun / kona]  powder, flour

┏━ 疒  (sickness 疒)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 痛  [tsū / ita(i)]  pain
┗━ 疲  [tsuka(reru)]  tired

┏━ 門  (gate 門)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 閉  [hei / shi(maru), shi(meru)]  close
┗━ 関  [kan]  relation, connection

┏━ 隹鳥  (bird 隹/鳥)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 難  [nan / muzuka(shii)]  difficult
┗━ 鳴  [na(ku), na(ru)]  cry, ring

┏━ 皿  (dish 皿)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 皿  [sara]  plate
┗━ 血  [ketsu / chi]  blood

┏━ 尸  (corpse/roof 尸)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 局  [kyoku]  bureau, office
┗━ 届  [todo(keru), todo(ku)]  deliver, reach

┏━ 彳  (step 彳)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 役  [yaku]  role, duty
┗━ 復  [fuku]  restore, again

┏━ 舟  (boat 舟)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 船  [sen / fune, funa]  boat, ship
┗━ 般  [han]  general

┏━ 角  (horn 角)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 角  [kaku / kado]  corner, angle
┗━ 解  [kai]  understand, solve

┏━ 冫  (ice 冫)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 冷  [rei / tsume(tai), hi(yasu), hi(eru), sa(meru), sa(masu)]  cold, cool
┗━ 凍  [tō / kō(ru)]  freeze

┏━ 金  (metal 金)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 鉄  [tetsu]  iron, steel
┗━ 録  [roku]  record

┏━ 戸  (door 戸)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 戻  [modo(ru), modo(su)]  return
┗━ 戸  [to]  door

┏━ 歯  (tooth 歯)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 齢  [rei]  age
┗━ 歯  [shi / ha]  tooth

┏━ 酉  (wine/chem 酉)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 配  [hai / kuba(ru)]  distribute, deliver
┗━ 酒  [shu / sake, saka]  alcohol

┏━ 戈  (halberd 戈)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 戦  [sen / tataka(u)]  war, fight
┗━ 成  [sei]  become, achieve

┏━ 厂  (cliff 厂)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 厚  [atsu(i)]  thick
┗━ 原  [gen]  original, source

┏━ 一  (one 一)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 両  [ryō]  both
┗━ 並  [nara(bu), nara(beru)]  line up

┏━ 化比  ('change/compare' 化・比)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 化  [ka/ke]  change, transform
┗━ 比  [hi / kura(beru)]  compare

┏━ 冂  (box 冂)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 再  [sai/sa]  again, re-
┗━ 冊  [satsu]  counter for books

┏━ 亅  (hook 亅)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 予  [yo]  beforehand, expectation
┗━ 了  [ryō]  finish, understand

┏━ 入  (enter 入)  [2 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┣━ 内  [nai / uchi]  inside
┗━ 全  [zen]  all, whole

┏━ 子  (child 子)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 存  [zon]  exist, know

┏━ 里  (village 里)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 量  [ryō]  amount, quantity

┏━ 工  (work 工)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 差  [sa / sa(su)]  difference, insert

┏━ 求  ('seek' 求 family)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 求  [kyū / moto(meru)]  seek, request

┏━ 見  (see 見)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 観  [kan]  observe, view

┏━ 羊  (sheep 羊)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 美  [bi / utsuku(shii)]  beautiful

┏━ 行  (road 行)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 術  [jutsu]  technique, art

┏━ 彡  (bristle 彡)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 形  [kei/gyō / katachi]  shape, form

┏━ 卩  (seal 卩)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 卵  [ran / tamago]  egg

┏━ 歹  (death-side 歹)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 残  [zan / noko(ru), noko(su)]  remain

┏━ 耳  (ear 耳)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 職  [shoku]  job, occupation

┏━ 車  (vehicle 車)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 輸  [yu]  transport, export/import

┏━ 夂  ('go/follow' 夂)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 変  [hen / ka(waru), ka(eru)]  change

┏━ 白  (white 白)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 的  [teki]  target, -al

┏━ 弋  ('style' 弋)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 式  [shiki]  style, ceremony

┏━ 斤  (axe 斤)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 断  [dan / kotowa(ru)]  refuse, cut off

┏━ 殳  (weapon 殳)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 段  [dan]  step, stage

┏━ 勹  (wrap 勹)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 危  [ki / abunai]  dangerous

┏━ 飛  ('fly' 飛)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 飛  [hi / to(bu)]  fly

┏━ 非  ('wrong' 非)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 非  [hi]  non-, not

┏━ 面  (face 面)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 面  [men / omote]  surface, face

┏━ 交  ('mix' 交)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 交  [kō]  mix, interact

┏━ 公  (public 八/公)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 公  [kō]  public

┏━ 鼻  (nose 鼻)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 鼻  [bi / hana]  nose

┏━ 骨  (bone 骨)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 骨  [kotsu / hone]  bone

┏━ 黄  (yellow 黄)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 黄  [ō / ki]  yellow

┏━ 乳  ('milk' 乚)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 乳  [nyū]  milk

┏━ 癶  (footsteps 癶)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 登  [tō/to / nobo(ru)]  climb, register

┏━ 支  (branch 支)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 支  [shi]  support, branch

┏━ 止  (stop 止)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 歳  [sai]  years old

┏━ 小  (small 小)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 当  [tō / a(taru)]  hit, appropriate

┏━ 厶  (private 厶)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 参  [san / mai(ru)]  participate, go/come

┏━ 丷  ('divide' 丷)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 単  [tan]  simple, single

┏━ 辛  (bitter 辛)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 辞  [ji / ya(meru)]  resign, dictionary

┏━ 身  (body 身)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 身  [shin / mi]  body, oneself

┏━ 虫  (insect 虫)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 虫  [mushi]  insect

┏━ 穴  (cave 穴)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 窓  [mado]  window

┏━ 山  (mountain 山)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 島  [tō / shima]  island

┏━ 西  (cover/west 覀)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 要  [yō / i(ru)]  need, necessary

┏━ 足  (foot 足)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 路  [ro]  road, route

┏━ 十  (ten 十)  [1 kanji] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┗━ 協  [kyō]  cooperate
"""

def parse_groups(raw_text):
    groups = []
    current_group = None
    for line in raw_text.strip().split('\n'):
        line = line.strip()
        if not line:
            continue
        if line.startswith('┏━'):
            m = re.search(r'┏━\s*(.*?)\s*\[(\d+)\s*kanji\]', line)
            if m:
                rad_title = m.group(1).strip()
                count = int(m.group(2))
                current_group = {
                    'title': rad_title,
                    'expected_count': count,
                    'items': []
                }
                groups.append(current_group)
        elif line.startswith('┣━') or line.startswith('┗━'):
            m = re.search(r'[┣┗]━\s*(\S+)\s*\[(.*?)\]\s*(.*)', line)
            if m:
                kanji = m.group(1).strip()
                reading = m.group(2).strip()
                meaning = m.group(3).strip()
                if current_group:
                    current_group['items'].append({
                        'kanji': kanji,
                        'reading': reading,
                        'meaning': meaning
                    })
    return groups

# Palette definitions
BG_COLOR = colors.HexColor('#FFFFFF')
HEADER_GREEN = colors.HexColor('#1E8449')
HEADER_TEXT_MAIN = colors.HexColor('#0F172A')
HEADER_TEXT_SUB = colors.HexColor('#475569')

RADICAL_BOX_BG = colors.HexColor('#1E293B')  # Dark slate header box for radical
RADICAL_BOX_TXT = colors.HexColor('#FFFFFF')

KANJI_BOX_BG = colors.HexColor('#F8FAFC')
KANJI_BOX_BORDER = colors.HexColor('#475569')
KANJI_CHAR_COLOR = colors.HexColor('#8C2A1E') # Dark Japanese crimson
MEANING_COLOR = colors.HexColor('#1E293B')
READING_COLOR = colors.HexColor('#4B5563')
LINE_COLOR = colors.HexColor('#94A3B8')

def draw_header_footer(c, page_num, total_pages):
    # Header Badge (Top Left)
    c.setFillColor(HEADER_GREEN)
    c.roundRect(36, 742, 110, 36, 4, stroke=0, fill=1)
    
    c.setFont('ArialUnicode', 11)
    c.setFillColor(colors.white)
    c.drawString(46, 762, "JLPT N3")
    c.setFont('ArialUnicode', 8.5)
    c.drawString(46, 749, "Kanji Mind Map")
    
    # Header Title Text (Top Middle/Right)
    c.setFillColor(HEADER_TEXT_MAIN)
    c.setFont('ArialUnicode', 13)
    c.drawString(158, 762, "JLPT N3 Kanji Mind Map — Weeks 1–6")
    
    c.setFillColor(HEADER_TEXT_SUB)
    c.setFont('ArialUnicode', 9)
    c.drawString(158, 748, "336 Kanji arranged into 104 Component Radical Trees")
    
    # Header Separator
    c.setStrokeColor(colors.HexColor('#CBD5E1'))
    c.setLineWidth(0.8)
    c.line(36, 734, 576, 734)
    
    # Footer Separator
    c.line(36, 36, 576, 36)
    
    # Footer Text
    c.setFont('ArialUnicode', 8)
    c.setFillColor(colors.HexColor('#64748B'))
    c.drawString(36, 22, "kanji60s.com • JLPT N3 Mind Map Flowcharts")
    c.drawRightString(576, 22, f"Page {page_num} of {total_pages}")

def draw_kanji_node(c, x, y, item):
    """
    Draws a single kanji node centered at (x, y).
    x, y is the center of the Kanji Box (28 x 28 pt).
    """
    kanji = item['kanji']
    reading = item['reading']
    meaning = item['meaning']
    
    # Box dimensions
    w_box, h_box = 28, 28
    x_box = x - w_box / 2.0
    y_box = y - h_box / 2.0
    
    # 1. Meaning text above box
    c.setFont('ArialUnicode', 7.5)
    c.setFillColor(MEANING_COLOR)
    # Truncate meaning if too long to prevent horizontal overlap
    disp_meaning = meaning if len(meaning) <= 15 else meaning[:13] + ".."
    c.drawCentredString(x, y_box + h_box + 4, disp_meaning)
    
    # 2. Kanji Box
    c.setFillColor(KANJI_BOX_BG)
    c.setStrokeColor(KANJI_BOX_BORDER)
    c.setLineWidth(0.75)
    c.roundRect(x_box, y_box, w_box, h_box, 3, stroke=1, fill=1)
    
    # Kanji Character
    c.setFont('ArialUnicode', 14)
    c.setFillColor(KANJI_CHAR_COLOR)
    c.drawCentredString(x, y_box + 7, kanji)
    
    # 3. Romaji Reading below box
    c.setFont('ArialUnicode', 6.5)
    c.setFillColor(READING_COLOR)
    disp_reading = f"[{reading}]"
    if len(disp_reading) > 16:
        disp_reading = disp_reading[:14] + "..]"
    c.drawCentredString(x, y_box - 10, disp_reading)
    
    return (x_box, y_box, w_box, h_box)

def calculate_group_layout(group):
    """
    Calculates relative layout coordinates for a group's tree.
    Returns: (num_cols, rows_per_col, total_vertical_rows, layout_data)
    """
    items = group['items']
    N = len(items)
    
    if N <= 7:
        num_cols = 1
    elif N <= 15:
        num_cols = 2
    else:
        num_cols = 3
        
    items_per_col = math.ceil(N / num_cols)
    return num_cols, items_per_col

def generate_pdf(output_filename="JLPT_N3_Kanji_Mind_Map.pdf"):
    groups = parse_groups(RAW_DATA)
    
    c = canvas.Canvas(output_filename, pagesize=letter)
    
    # Page setup
    margin_x = 36
    top_y = 715
    bottom_y = 50
    page_height_available = top_y - bottom_y  # ~665 pt
    
    # Grouping onto pages
    row_height = 48  # vertical pitch per kanji node
    group_padding = 24  # gap between radical groups
    
    pages_data = []
    current_page_groups = []
    current_page_height = 0
    
    for g in groups:
        num_cols, rows_per_col = calculate_group_layout(g)
        g_height = max(rows_per_col * row_height + 10, 50)
        
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
    print(f"Total pages generated: {total_pages}")
    
    # Render pages
    for page_idx, page_groups in enumerate(pages_data):
        page_num = page_idx + 1
        draw_header_footer(c, page_num, total_pages)
        
        curr_y = top_y - 10
        
        for g, g_height, num_cols, rows_per_col in page_groups:
            items = g['items']
            N = len(items)
            
            # Group top y and bottom y
            group_top_y = curr_y
            group_center_y = group_top_y - (g_height / 2.0)
            
            # Draw Radical Root Box
            x_root = 42
            w_root = 100
            h_root = 32
            y_root = group_center_y - (h_root / 2.0)
            
            c.setFillColor(RADICAL_BOX_BG)
            c.setStrokeColor(colors.HexColor('#0F172A'))
            c.setLineWidth(0.8)
            c.roundRect(x_root, y_root, w_root, h_root, 4, stroke=1, fill=1)
            
            # Text inside radical box
            c.setFillColor(RADICAL_BOX_TXT)
            c.setFont('ArialUnicode', 9)
            c.drawCentredString(x_root + w_root/2.0, y_root + 18, g['title'])
            c.setFont('ArialUnicode', 7.5)
            c.setFillColor(colors.HexColor('#94A3B8'))
            c.drawCentredString(x_root + w_root/2.0, y_root + 6, f"[{N} kanji]")
            
            root_connect_x = x_root + w_root
            root_connect_y = group_center_y
            
            # Column X positions
            if num_cols == 1:
                col_x_offsets = [210]
            elif num_cols == 2:
                col_x_offsets = [200, 370]
            else:
                col_x_offsets = [185, 315, 445]
                
            # Place items into columns
            idx = 0
            for col in range(num_cols):
                x_node = col_x_offsets[col]
                col_items = items[col * rows_per_col : (col + 1) * rows_per_col]
                
                start_y_col = group_top_y - 24
                
                for r_idx, item in enumerate(col_items):
                    y_node = start_y_col - (r_idx * row_height)
                    
                    # Draw connecting line from root box to kanji node
                    c.setStrokeColor(LINE_COLOR)
                    c.setLineWidth(0.9)
                    child_connect_x = x_node - 14  # left edge of box
                    c.line(root_connect_x, root_connect_y, child_connect_x, y_node)
                    
                    # Draw Kanji Node
                    draw_kanji_node(c, x_node, y_node, item)
                    idx += 1
                    
            curr_y -= (g_height + group_padding)
            
        c.showPage()
        
    c.save()
    print(f"Successfully generated {output_filename}!")

if __name__ == "__main__":
    generate_pdf("kflow.pdf")

"""第5章: 請求書テンプレートのセル配置と、各セルにどのシートの何を書くか（PART-WHOLE + SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(1000, 610)
h = 32

today = c.node(70, 88, "今日", emoji_cp="1f4c5", w=60, h=50, label_scale="label")   # 📅

c.text(30, 170, "顧客シート", scale="label", align="left")
cy0 = 186
table(c, 30, cy0, [
    ["顧客名", "郵便番号", "住所", "ご担当者"],
    ["ひまわり商事", "〒104-0000", "東京都中央区…", "総務部 山田 様"],
    ["あおば設計事務所", "〒160-0000", "東京都新宿区…", "事務担当 佐藤 様"],
], [120, 90, 100, 110], color="blue", h=h, fills={(1, k): "yellow" for k in range(4)})

c.text(30, 346, "売上シート", scale="label", align="left")
sy0 = 362
table(c, 30, sy0, [
    ["日付", "顧客", "商品", "単価", "数量", "金額"],
    ["10/02", "あおば設計事務所", "付箋", "250", "5", "1250"],
    ["10/07", "ひまわり商事", "ホッチキス", "680", "2", "1360"],
    ["10/07", "ひまわり商事", "コピー用紙", "550", "3", "1650"],
    ["10/07", "みなと物産", "ノート", "180", "10", "1800"],
    ["…", "…", "…", "…", "…", "…"],
], [56, 120, 84, 50, 46, 64], color="blue", h=h,
    fills={(r, k): "yellow" for r in (2, 3) for k in range(6)})

# 右: 請求書テンプレート
c.sticky(590, 20, 390, 570, color="gray", weight="thin", rx=8)
c.text(785, 60, "請求書テンプレート", scale="label")
e3 = c.sticky(800, 76, 160, 44, color="yellow", rx=6, weight="thin")
c.text(880, 105, "E3 発行日", scale="caption")
addr = c.sticky(610, 150, 270, 110, color="yellow", rx=6, weight="thin")
c.text(745, 186, "A5 宛名（〇〇 御中）", scale="caption")
c.text(745, 218, "A6 郵便番号と住所", scale="caption")
c.text(745, 248, "A7 ご担当者", scale="caption")
c.sticky(610, 290, 330, 46, color="orange", rx=6, weight="thin")
c.text(775, 320, "B9 ご請求金額（E33 と同じ）", scale="caption")
detail = c.sticky(610, 403, 350, 110, color="blue", rx=6, weight="thin")
c.text(785, 447, "A12:E31 明細", scale="label")
c.text(785, 481, "日付・商品・単価・数量・金額", scale="caption")
c.sticky(610, 528, 350, 46, color="orange", rx=6, weight="thin")
c.text(785, 558, "E33 合計（明細の金額を合計）", scale="caption")

c.link(today, (e3, "w"), label="Date", primary=False)
c.link((450, cy0 + h * 1.5), (addr, "w"), label="1 行")
c.link((450, sy0 + h * 3), (610, sy0 + h * 3), label="この顧客の行", label_pos="above")
c.save("01-template.svg")

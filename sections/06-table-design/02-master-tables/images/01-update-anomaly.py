"""6-2: 住所を全部の行に書くと、移転のたびに全部を直して漏れが出る。マスタなら 1 か所（BALANCE / SCALE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(1040, 380)
c.text(40, 60, "✕ 売上の行ごとに住所を書く", scale="heading", align="left")
table(c, 40, 90, [
    ["日付", "顧客", "住所"],
    ["10/07", "ひまわり商事", "港区はなみずき 1-1"],
    ["10/13", "ひまわり商事", "港区はなみずき 1-1"],
    ["10/15", "ひまわり商事", "中央区ひまわり 2-4-6"],
    ["10/21", "ひまわり商事", "港区はなみずき 1-1"],
    ["10/23", "ひまわり商事", "港区はなみずき 1-1"],
], [70, 130, 190], color="orange", h=34, fills={(3, 2): "red"})
c.text(235, 330, "移転したら全部の行を直す。1 行でも漏れると\n請求書の住所が古いまま", scale="caption")

c.text(600, 60, "◯ 顧客シートに 1 回だけ書く", scale="heading", align="left")
table(c, 600, 158, [
    ["顧客名", "住所"],
    ["ひまわり商事", "港区はなみずき 1-1"],
], [140, 210], color="blue", h=34, fills={(1, 1): "green"})
c.text(775, 260, "直すのは 1 セルだけ。\n売上の行はどれも、ここを引けば新しい住所になる", scale="caption")
c.save("01-update-anomaly.svg")

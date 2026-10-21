"""6-1: 見やすく整えた表の、集計やマクロを止める 4 つの崩れ（BLOCKAGE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(1030, 470)
x0, y0, h = 310, 90, 40
widths = [130, 100, 90, 70]
table(c, x0, y0, [
    ["顧客", "商品", "10月", "11月"],
    ["", "付箋", "5", "3"],
    ["", "ノート", "2 ※返品1", "4"],
    [("小計", 2), "7", "7"],
    [("", 4)],
    ["", "付箋", "10個", "6個"],
    ["", "ホッチキス", "1個", "2個"],
], widths, color="orange", h=h,
    fills={(3, 0): "yellow", (3, 2): "yellow", (3, 3): "yellow", (2, 2): "red"},
    vmerges=[(1, 0, 2, "ひまわり商事"), (5, 0, 2, "みなと物産")])

def note(x, y, title, body, color):
    n = c.sticky(x, y, 250, 84, color=color, weight="thin")
    c.text(x + 16, y + 32, title, scale="label", align="left")
    c.text(x + 16, y + 60, body, scale="caption", align="left")
    return n

n1 = note(24, 100, "① セルの結合", "2 行目の顧客が空になり\nSUMIF もフィルターも外れる", "red")
n2 = note(24, 300, "② 小計と空行", "合計に小計が混ざる\n範囲が空行で切れる", "red")
n3 = note(756, 200, "③ 数とメモが同じセル", "「2 ※返品1」は\n数として足せない", "red")
n4 = note(756, 60, "④ 月が横に並ぶ", "月が増えるたびに\n列も式もコードも直す", "red")
n5 = note(756, 340, "⑤ 値に単位", "「10個」は文字なので\n足し算に入らない", "red")
c.link(n1, (x0 + 30, y0 + h * 1.5), primary=False)
c.link(n2, (x0 + 30, y0 + h * 3.5), primary=False)
c.link(n3, (x0 + 316, y0 + h * 2.5), primary=False)
c.link(n4, (x0 + 382, y0 + h * 0.5), primary=False)
c.link(n5, (x0 + 382, y0 + h * 5.5), primary=False)
c.save("01-broken-table.svg")

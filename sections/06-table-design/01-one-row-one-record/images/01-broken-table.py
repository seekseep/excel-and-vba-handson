"""6-1: 見やすく整えた表の、集計やマクロを止める 5 つの崩れと、それぞれ何が困るか（BLOCKAGE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(1070, 490)
x0, y0, h = 340, 90, 40
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
    n = c.sticky(x, y, 280, 48 + 28 * (body.count("\n") + 1) + 16, color=color, weight="thin")
    c.text(x + 16, y + 32, title, scale="label", align="left")
    c.text(x + 16, y + 62, body, scale="caption", align="left")
    return n

n1 = note(14, 40, "① セルの結合", "値は先頭の 1 行だけ。\n2 行目の顧客が空で\nSUMIF・フィルター・\n並べ替えから漏れる", "red")
n2 = note(14, 300, "② 小計と空行", "列全体の合計に\n小計が混ざって 2 倍に。\nピボットは空行を\n表の終わりと見なす", "red")
n4 = note(776, 10, "④ 月が横に並ぶ", "月が増えるたびに\n列を足し、式も\nマクロの範囲も直す", "red")
n3 = note(776, 172, "③ 数とメモが混在", "「2 ※返品1」は文字で\n合計に入らない。\n返品の数も読めない", "red")
n5 = note(776, 334, "⑤ 値に単位", "「10個」は文字。\n見た目は数でも\n足し算に入らない", "red")
c.link(n1, (x0 + 30, y0 + h * 1.5), primary=False)
c.link(n2, (x0 + 30, y0 + h * 3.5), primary=False)
c.link(n3, (x0 + 316, y0 + h * 2.5), primary=False)
c.link(n4, (x0 + 382, y0 + h * 0.5), primary=False)
c.link(n5, (x0 + 382, y0 + h * 5.5), primary=False)
c.save("01-broken-table.svg")

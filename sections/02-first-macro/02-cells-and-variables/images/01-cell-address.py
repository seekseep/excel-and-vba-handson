"""セルの指し方。B2 は番地なら Range("B2")、番号なら Cells(2, 2)（行が先、列が後）（LINK）。

0-3 と 2-2 に同じ図を置く。直すときは両方の images/ で走らせる。
"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(1040, 520)
h = 48
widths = [56, 96, 96, 96, 96]
gx, gy = 362, 70
table(c, gx, gy, [
    ["", "A", "B", "C", "D"],
    ["1", "", "", "", ""],
    ["2", "", "", "", ""],
    ["3", "", "", "", ""],
    ["4", "", "", "", ""],
], widths, color="gray", h=h, scale="body",
    fills={(r, 0): "gray" for r in range(5)} | {(0, 2): "blue", (2, 0): "blue"})
b2 = c.sticky(gx + 152, gy + h * 2, 96, h, color="yellow", rx=2)
c.text(gx + 200, gy + h * 2 + h / 2 + 7, "B2", scale="label", fill="#0f172a")
c.text(gx + 200, 56, "2 列目（B）", scale="caption", fill="#1d4ed8")
c.text(gx - 12, gy + h * 2 + h / 2 + 5, "2 行目", scale="caption", fill="#1d4ed8", align="right")

addr = c.sticky(40, 350, 290, 110, color="blue", weight="thin")
c.text(185, 386, "番地で指す", scale="label")
c.text(185, 418, 'Range("B2")', scale="body", font="technical", fill="#0f172a")
c.text(185, 444, "getRange('B2')", scale="caption", font="technical")

num = c.sticky(710, 350, 290, 110, color="green", weight="thin")
c.text(855, 386, "行, 列の番号で指す", scale="label")
c.text(855, 418, "Cells(2, 2)", scale="body", font="technical", fill="#0f172a")
c.text(855, 444, "getRange(2, 2)", scale="caption", font="technical")

c.link((addr, "e"), (b2, "s"), route="elbow")
c.link((num, "n"), (b2, "e"), route="elbow")
c.text(520, 500, "Cells(2, 2) は（2 行目, 2 列目）。行が先、列が後", scale="label", fill="#0f172a")
c.save("01-cell-address.svg")

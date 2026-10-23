"""データのある最後の行。A 列のいちばん下（Rows.Count）から上へ飛ぶと、最後の明細の 31 行目に止まる（VERTICALITY + SOURCE-PATH-GOAL）。

0-3 と 3-1（91-last-row）に同じ図を置く。直すときは両方の images/ で走らせる。
"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(780, 460)
h = 40
gx, gy = 30, 30
widths = [90, 140]
rows = [
    ["", "A"],
    ["1", "日付"],
    ["2", "2026/10/02"],
    ["3", "2026/10/06"],
    ["⋮", "⋮"],
    ["30", "2026/10/28"],
    ["31", "2026/10/29"],
    ["32", ""],
    ["⋮", ""],
    ["1048576", ""],
]
fills = {(r, 0): "gray" for r in range(len(rows))}
fills.update({(6, 0): "yellow", (6, 1): "yellow", (9, 0): "blue", (9, 1): "blue"})
table(c, gx, gy, rows, widths, color="gray", h=h, scale="label", fills=fills)

right = gx + sum(widths)
y31 = gy + h * 6.5
ybottom = gy + h * 9.5
c.link((right + 8, ybottom), (right + 8, y31), route="elbow", start="e", end="e", stub=40)
c.text(right + 24, (y31 + ybottom) / 2 - 4, "End(xlUp)\nCtrl + ↑ で上へ", scale="caption", align="left")

tx = 430
c.text(tx, gy + h * 1.5, "Cells(Rows.Count, 1)", scale="body", align="left", font="technical", fill="#0f172a")
c.text(tx + 24, gy + h * 1.5 + 30, ".End(xlUp).Row", scale="body", align="left", font="technical", fill="#0f172a")
c.text(tx, gy + h * 1.5 + 64, "GAS は sheet.getLastRow()", scale="caption", align="left")
c.text(tx, y31 + 8, "lastRow = 31", scale="heading", align="left", fill="#b45309")
c.text(tx, y31 + 40, "値がある最後の行で止まる", scale="caption", align="left")
c.text(tx, ybottom + 6, "Rows.Count … いちばん下の行", scale="caption", align="left")
c.link((tx - 12, y31), (right + 24, y31), primary=False, dash="dashed")
c.link((tx - 12, ybottom), (right + 24, ybottom), primary=False, dash="dashed", head=False)
c.save("02-last-row.svg")

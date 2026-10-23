"""3-1: For のくり返し。i が 2, 3, 4 … 31 と変わり、1 周ごとに i 行目の D × E を F に書く（CYCLE / SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(1000, 410)
c.text(500, 44, "1 周で 1 行。i が行番号になる", scale="heading")

code = c.sticky(30, 110, 390, 190, color="gray", weight="thin")
c.text(50, 150, "For i = 2 To 31", scale="label", align="left", font="technical")
c.text(78, 196, "Cells(i, 6) =", scale="body", align="left", font="technical", fill="#0f172a")
c.text(100, 228, "Cells(i, 4) * Cells(i, 5)", scale="body", align="left", font="technical", fill="#0f172a")
c.text(50, 276, "Next i", scale="label", align="left", font="technical")
c.text(225, 344, "i を 1 ずつ増やして 31 まで 30 回くり返す", scale="body")
c.text(225, 374, "GAS では for (let i = 2; i <= 31; i++) { }", scale="caption")

H = 40
ROWS = [("2", "250", "5", "1250"), ("3", "250", "10", "2500"), ("4", "680", "2", "1360"),
        ("⋮", "⋮", "⋮", "⋮"), ("31", "60", "20", "1200")]
TX, TY = 620, 90
table(c, TX, TY, [["行", "D 単価", "E 数量", "F 金額"], *ROWS], [60, 90, 90, 110],
      color="blue", h=H, scale="body",
      fills={(r, 3): "orange" for r in (1, 2, 3, 5)})
c.text(525, TY + 20, "i の値", scale="body")
c.text(TX + 175, TY + H * 6 + 34, "どの行も D × E を F に書く", scale="body")

for r, (row, *_rest) in enumerate(ROWS, start=1):
    if row == "⋮":
        continue
    y = TY + r * H + H / 2
    chip = c.sticky(470, y - 15, 110, 30, color="yellow", rx=15, weight="thin")
    c.text(525, y + 6, f"i = {row}", scale="body")
    c.link(chip, (TX - 4, y), primary=False)
c.save("03-for-loop.svg")

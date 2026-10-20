"""第4章: Excel のオブジェクトの入れ子。Application ⊃ Workbook ⊃ Worksheet ⊃ Range（CONTAINER）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(900, 460)
c.sticky(40, 40, 820, 380, color="gray", weight="thin")
c.text(70, 78, "Excel（Application）", scale="label", align="left")
c.sticky(80, 100, 740, 290, color="green", weight="thin")
c.text(110, 138, "ブック（Workbook） ― 例: ThisWorkbook", scale="label", align="left")
c.sticky(120, 160, 660, 200, color="blue", weight="thin")
c.text(150, 198, "シート（Worksheet） ― 例: Worksheets(\"売上\")", scale="label", align="left")
c.sticky(160, 220, 580, 110, color="yellow", weight="thin")
c.text(190, 258, "セル（Range） ― 例: Range(\"B2\")、Cells(2, 2)", scale="label", align="left")
c.text(190, 300, ".Value で値、.Interior.Color で背景色", scale="caption", align="left")
c.save("00-thumbnail.svg")

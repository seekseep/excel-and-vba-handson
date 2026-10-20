"""第1章: 月末作業の流れ。前半（集計）は関数で、後半（シート作成・PDF）は VBA の出番（SOURCE-PATH-GOAL + CONTAINER）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1000, 340)
c.sticky(40, 60, 420, 230, color="green", weight="thin", dash="dashed")
c.text(250, 92, "Excel 関数でできる", scale="label")
c.sticky(540, 60, 420, 230, color="blue", weight="thin", dash="dashed")
c.text(750, 92, "VBA の出番", scale="label")

sales = c.node(140, 190, "売上明細\n30 行", emoji_cp="1f4ca")       # 📊
total = c.node(360, 190, "顧客ごとの\n合計", emoji_cp="1f9ee")      # 🧮
invoice = c.node(640, 190, "請求書\n5 社分", emoji_cp="1f4c4")      # 📄
pdf = c.node(860, 190, "PDF で\n保存", emoji_cp="1f4c1")           # 📁
c.link(sales, total)
c.link(total, invoice)
c.link(invoice, pdf)
c.save("01-monthly-work.svg")

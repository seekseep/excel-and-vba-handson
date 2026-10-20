"""第5章サムネ: ボタンを押す → マクロが動く → 請求書シートと PDF ができる（SOURCE-PATH-GOAL + SPLITTING）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(920, 340)
button = c.node(130, 170, "ボタンを押す", emoji_cp="1f446")              # 👆
macro = c.node(420, 170, "CreateInvoices", emoji_cp="2699")              # ⚙️
sheets = c.node(750, 80, "請求書シート × 5", emoji_cp="1f4c4")          # 📄
pdf = c.node(750, 260, "PDF × 5", emoji_cp="1f4c1")                     # 📁
c.link(button, macro, label="イベント")
c.link(macro, sheets)
c.link(macro, pdf)
c.save("00-thumbnail.svg")

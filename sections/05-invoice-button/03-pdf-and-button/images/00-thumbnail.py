"""5-3 サムネ: ボタンを押す → マクロが動く → 請求書シートと PDF ができる（SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(880, 260)
button = c.node(130, 120, "ボタンを押す", emoji_cp="1f446")          # 👆
macro = c.node(440, 120, "CreateInvoices", emoji_cp="2699")          # ⚙️
pdf = c.node(750, 120, "PDF × 5", emoji_cp="1f4c1")                  # 📁
c.link(button, macro, label="イベント")
c.link(macro, pdf, label="書き出す")
c.save("00-thumbnail.svg")

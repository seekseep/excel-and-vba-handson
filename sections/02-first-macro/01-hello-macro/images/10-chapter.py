"""2 章で作るもの: メッセージを出す → セルに書く → 読んで計算して書く、の 3 つの小さなマクロ（SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1000, 300)
a = c.node(160, 130, "メッセージを出す", emoji_cp="1f4ac", label_scale="label")       # 💬
b = c.node(500, 130, "セルに書く", emoji_cp="270f", label_scale="label")              # ✏️
d = c.node(840, 130, "読んで計算して書く", emoji_cp="1f9ee", label_scale="label")     # 🧮
c.text(160, a.below(26), "Hello", scale="caption")
c.text(500, b.below(26), "WriteCells", scale="caption")
c.text(840, d.below(26), "CalcTaxIncluded（B1 → B2）", scale="caption")
c.link(a, b)
c.link(b, d)
c.save("10-chapter.svg")

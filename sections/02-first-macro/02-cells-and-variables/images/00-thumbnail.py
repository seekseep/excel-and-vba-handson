"""2-2 サムネ: セルから読んで、変数に入れて、計算して、セルに書く（SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(880, 260)
b1 = c.node(130, 120, "B1  1200", emoji_cp="1f4ca")         # 📊
price = c.node(440, 120, "price", emoji_cp="1f4e6")         # 📦 変数＝箱
b2 = c.node(750, 120, "B2  1320", emoji_cp="1f4ca")         # 📊
c.link(b1, price, label="読む")
c.link(price, b2, label="× 1.1 して書く")
c.save("00-thumbnail.svg")

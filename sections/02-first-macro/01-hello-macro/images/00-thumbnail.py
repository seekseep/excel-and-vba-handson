"""第2章サムネ: コードを書く → 実行する → セルに結果が出る（SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(880, 300)
c.text(440, 52, "書いて、動かして、セルに出す", scale="heading")
code = c.node(150, 170, "Sub を書く", emoji_cp="1f4dd")      # 📝
run = c.node(440, 170, "F5 で実行", emoji_cp="25b6")          # ▶️
cell = c.node(730, 170, "B2 に 1320", emoji_cp="1f4ca")      # 📊
c.link(code, run)
c.link(run, cell, label="Range(\"B2\")")
c.save("00-thumbnail.svg")

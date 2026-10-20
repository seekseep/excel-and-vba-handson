"""3-4 サムネ: 合わない合計 → 止めて変数を見る → 原因を直すと一致する（SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(900, 260)
bug = c.node(130, 120, "67,330（合わない）", emoji_cp="1f41e")       # 🐞
look = c.node(450, 120, "i = 3 から始まっている", emoji_cp="1f50d")   # 🔍
ok = c.node(770, 120, "68,580（一致）", emoji_cp="2705")             # ✅
c.link(bug, look, label="止めて見る")
c.link(look, ok, label="直す")
c.save("00-thumbnail.svg")

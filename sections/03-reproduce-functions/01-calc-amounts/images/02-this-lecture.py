"""3-1: この節の流れ。関数で答えを用意してから、各行の金額を VBA で出して見比べる（SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1000, 400)
c.text(500, 44, "この節の流れ：答えを用意してから、VBA で出す", scale="heading")

rows = c.node(110, 225, "売上明細\n単価・数量 30 行", emoji_cp="1f4d2")        # 📒

func = c.sticky(240, 130, 240, 190, color="yellow", weight="thin")
c.text(360, 170, "① 関数で答えを用意", scale="label")
c.text(360, 232, "F2 に =D2*E2\nF31 までコピー", scale="body")
c.text(360, 290, "（SUM・SUMIF も）", scale="body")

vba = c.sticky(530, 130, 240, 190, color="blue", weight="thin")
c.text(650, 170, "② VBA で金額を出す", scale="label")
c.text(650, 232, "単価 × 数量 を\n31 行目までくり返す", scale="body")
c.text(650, 290, "（For 〜 Next）", scale="body")

ok = c.node(890, 225, "①と②の金額が\n同じか見比べる", emoji_cp="2705")          # ✅

c.link(rows, func)
c.link(func, vba)
c.link(vba, ok)
c.save("02-this-lecture.svg")

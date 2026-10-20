"""第1章サムネ: 手作業 → 関数 → VBA と道具を持ち替えて、月末の請求書づくりを片付ける（SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(880, 300)
c.text(440, 52, "月末の請求書づくり、どの道具で片付ける？", scale="heading")
manual = c.node(150, 170, "手作業", emoji_cp="270f")      # ✏️
func = c.node(440, 170, "Excel 関数", emoji_cp="1f9ee")   # 🧮
vba = c.node(730, 170, "VBA", emoji_cp="2699")            # ⚙️
c.link(manual, func, label="集計を任せる")
c.link(func, vba, label="操作を任せる")
c.save("00-thumbnail.svg")

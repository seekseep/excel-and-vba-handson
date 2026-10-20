"""第3章サムネ: 同じ売上明細を関数と VBA の 2 通りで集計し、答えが一致する（SPLITTING → 合流）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(900, 360)
sales = c.node(130, 180, "売上明細", emoji_cp="1f4ca")              # 📊
func = c.node(450, 85, "関数  =SUM", emoji_cp="1f9ee")             # 🧮
vba = c.node(450, 275, "VBA  For 〜 Next", emoji_cp="2699")        # ⚙️
ok = c.node(770, 180, "どちらも 68,580", emoji_cp="2705")          # ✅
c.link(sales, func)
c.link(sales, vba)
c.link(func, ok)
c.link(vba, ok)
c.save("00-thumbnail.svg")

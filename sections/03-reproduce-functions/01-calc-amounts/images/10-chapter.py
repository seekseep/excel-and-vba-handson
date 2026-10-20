"""3 章で作るもの: 売上明細から、金額 → 売上合計 → 顧客別の合計を VBA で出し、関数の答えと見比べる（SOURCE-PATH-GOAL + CONTAINER）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1160, 360)
c.sticky(250, 40, 880, 290, color="blue", weight="thin", dash="dashed")
c.text(690, 76, "関数と同じ答えを VBA で出す", scale="label")
sales = c.node(120, 180, "売上明細 30 行", emoji_cp="1f4ca", label_scale="label")   # 📊
amt = c.node(400, 180, "金額", emoji_cp="1f9ee", label_scale="label")                # 🧮
tot = c.node(690, 180, "売上合計", emoji_cp="1f4b0", label_scale="label")            # 💰
cus = c.node(980, 180, "顧客別の合計", emoji_cp="1f465", label_scale="label")        # 👥
c.text(400, amt.below(26), "F 列（For でくり返す）", scale="caption")
c.text(690, tot.below(26), "J2 = 68,580（SUM）", scale="caption")
c.text(980, cus.below(26), "J5〜J9（If で選ぶ・SUMIF）", scale="caption")
c.link(sales, amt)
c.link(amt, tot)
c.link(tot, cus)
c.save("10-chapter.svg")

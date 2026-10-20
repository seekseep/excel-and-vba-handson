"""3-2 サムネ: 明細を 1 行ずつ見て、条件に合う行だけ合計に足す（SOURCE-PATH-GOAL + BLOCKAGE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(900, 300)
rows = c.node(130, 140, "明細を 1 行ずつ", emoji_cp="1f4ca")         # 📊
judge = c.node(450, 140, "顧客名が同じ？", shape="diamond", color="yellow", w=200, h=110, label_scale="label")
total = c.node(770, 140, "total に足す", emoji_cp="1f9ee")           # 🧮
c.link(rows, judge, label="For")
c.link(judge, total, label="If")
c.save("00-thumbnail.svg")

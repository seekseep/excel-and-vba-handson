"""6-3: 表をきちんと設計するかどうかを 3 つの問いで決める（SPLITTING）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(900, 380)
qs = c.sticky(30, 40, 330, 300, color="yellow", weight="thin", rx=12)
c.text(195, 84, "3 つの問い", scale="label")
for k, q in enumerate(("毎月続く？", "人に渡す？", "関数やマクロで集計したい？")):
    c.text(56, 150 + k * 64, f"{k + 1}. {q}", scale="body", align="left")

yes = c.sticky(530, 40, 340, 130, color="green", weight="thin", rx=12)
c.text(700, 96, "最初に 5 分かけて", scale="caption")
c.text(700, 132, "1 件と列を決める", scale="heading")
no = c.sticky(530, 210, 340, 130, color="gray", weight="thin", rx=12)
c.text(700, 262, "そのままでよい", scale="heading")
c.text(700, 300, "一度きりの表、自分だけの数十行", scale="caption")

c.link((qs, "e"), (yes, "w"), route="elbow")
c.link((qs, "e"), (no, "w"), route="elbow", primary=False)
c.text(445, 80, "どれか はい", scale="label")
c.text(445, 318, "全部 いいえ", scale="label")
c.save("03-when-to-design.svg")

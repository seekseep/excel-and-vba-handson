"""1-1: 月末の 5 つの作業。1〜3 の集計は関数でできて、4・5 のシート作成と PDF は VBA の出番（SOURCE-PATH-GOAL + CONTAINER）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1240, 420)
c.sticky(24, 40, 700, 340, color="green", weight="thin", dash="dashed")
c.text(374, 76, "Excel 関数でできる", scale="label")
c.sticky(752, 40, 464, 340, color="blue", weight="thin", dash="dashed")
c.text(984, 76, "VBA の出番", scale="label")

steps = [
    (140, "1. 金額を出す", "単価 × 数量\nを各行に", "1f9ee"),        # 🧮
    (374, "2. 売上合計", "月の合計を\n出す", "1f4ca"),                # 📊
    (608, "3. 顧客ごとの合計", "5 社それぞれの\n合計を出す", "1f465"),  # 👥
    (868, "4. 請求書を作る", "ひな形をコピーして\n宛名と明細を書き写す", "1f4c4"),  # 📄
    (1100, "5. PDF で保存", "5 社分を\nそれぞれ保存", "1f4c1"),      # 📁
]
nodes = []
for x, title, detail, cp in steps:
    n = c.node(x, 190, title, emoji_cp=cp, label_scale="label")
    c.text(x, n.below(26), detail, scale="caption")
    nodes.append(n)
for a, b in zip(nodes, nodes[1:]):
    c.link(a, b)
c.save("01-monthly-work.svg")

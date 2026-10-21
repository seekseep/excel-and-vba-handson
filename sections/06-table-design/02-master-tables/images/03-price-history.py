"""6-2: 単価はマスタから引くだけにすると、値上げで過去の請求額まで変わる。明細に書き写して残す（COUNTERFORCE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(1000, 420)
master = table(c, 330, 40, [
    ["商品", "単価"],
    ["付箋", "250 → 280"],
], [120, 220], color="blue", h=36, fills={(1, 1): "yellow"})
c.text(690, 86, "← 11 月 1 日に付箋を値上げした", scale="caption", align="left")

c.text(220, 312, "✕ 明細はマスタの単価を引くだけ", scale="label")
bad = table(c, 40, 200, [
    ["日付", "商品", "数量", "単価", "金額"],
    ["10/02", "付箋", "5", "280", "1,400"],
], [70, 80, 60, 70, 80], color="orange", h=36, fills={(1, 3): "red", (1, 4): "red"})
c.text(220, 342, "10 月分まで 280 円で計算し直される。\n送った請求書（1,250 円）と合わない", scale="caption")

c.text(740, 312, "◯ 売った時点の単価を明細に書き写す", scale="label")
good = table(c, 560, 200, [
    ["日付", "商品", "数量", "単価", "金額"],
    ["10/02", "付箋", "5", "250", "1,250"],
], [70, 80, 60, 70, 80], color="green", h=36, fills={(1, 3): "green", (1, 4): "green"})
c.text(740, 342, "値上げしても 10 月の明細は 250 円のまま。\nマスタは「これから売る値段」として使う", scale="caption")
c.link(master, bad, primary=False, dash="dashed", label="引く")
c.link(master, good, primary=False, label="書き写す")
c.save("03-price-history.svg")

"""4 章で作るもの: 売上シートを顧客ごとのシートに振り分け、ブックのコピーを保存する（関数ではできない操作）（SPLITTING + CONTAINER）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1040, 420)
sales = c.node(140, 210, "売上シート", emoji_cp="1f4ca", label_scale="label")        # 📊
group = c.sticky(400, 30, 300, 360, color="gray", weight="thin", dash="dashed")
c.text(550, 62, "顧客ごとのシート", scale="label")
names = ["ひまわり商事", "あおば設計事務所", "みなと物産", "さくら学習塾", "こもれび珈琲"]
for i, n in enumerate(names):
    y = 82 + i * 58
    c.sticky(440, y, 220, 44, color="yellow", rx=8, weight="thin")
    c.text(550, y + 29, n, scale="body")
save = c.node(900, 210, "ブックのコピーを保存", emoji_cp="1f4be", label_scale="label")  # 💾
c.link(sales, group, label="振り分ける")
c.link(group, save, label="保存")
c.save("10-chapter.svg")

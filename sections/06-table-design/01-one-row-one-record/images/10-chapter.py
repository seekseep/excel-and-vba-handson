"""第 6 章: 1 行に 1 件の表にする → 同じことをマスタに分ける → 新しいシートを設計する（SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1000, 300)
steps = [
    (170, "6-1 1 行に 1 件", "見やすさより\n並べ方をそろえる", "1f4c3"),   # 📃
    (500, "6-2 マスタに分ける", "同じことは\n1 か所にだけ書く", "1f5c2"),  # 🗂️
    (830, "6-3 新しいシートを作る", "1 件を決めてから\n列を決める", "1f4d0"),  # 📐
]
nodes = []
for x, title, detail, cp in steps:
    n = c.node(x, 120, title, emoji_cp=cp, label_scale="label")
    c.text(x, n.below(26), detail, scale="caption")
    nodes.append(n)
for a, b in zip(nodes, nodes[1:]):
    c.link(a, b)
c.save("10-chapter.svg")

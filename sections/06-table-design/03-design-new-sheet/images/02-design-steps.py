"""6-3: 新しいシートを作るときの 5 つの手順（SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1160, 300)
steps = [
    (110, "1. 見せる表を描く", "最後に何を見たいか", "1f3af"),        # 🎯
    (345, "2. 1 件を決める", "1 行は何の 1 回か", "1f9e9"),           # 🧩
    (580, "3. 列を決める", "1 列に 1 種類\n1 セルに 1 つの値", "1f4cf"),  # 📏
    (815, "4. マスタに分ける", "何度も出てくるものに\nコードを振る", "1f5c2"),  # 🗂️
    (1050, "5. 入力を助ける", "テーブル・プルダウン\n日付は日付で", "2705"),  # ✅
]
nodes = []
for x, title, detail, cp in steps:
    n = c.node(x, 120, title, emoji_cp=cp, label_scale="label")
    c.text(x, n.below(26), detail, scale="caption")
    nodes.append(n)
for a, b in zip(nodes, nodes[1:]):
    c.link(a, b)
c.save("02-design-steps.svg")

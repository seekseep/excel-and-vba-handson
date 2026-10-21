"""6-1: データの表の 3 つの決まり。1 行に 1 件、1 列に 1 種類、1 セルに 1 つの値（PART-WHOLE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1000, 330)
cards = [
    (40, "1 行に 1 件", "1f4c3", "1 行 = 1 回の売上\n小計・空行・メモは入れない\n見出しは 1 行だけ"),   # 📃
    (360, "1 列に 1 種類", "1f4cf", "日付の列には日付だけ\n月や顧客を列に並べない\n単位は値でなく見出しに"),  # 📏
    (680, "1 セルに 1 つの値", "1f9e9", "「付箋×5」→ 商品と数量に\n数とメモを分ける\nセルを結合しない"),       # 🧩
]
for x, title, cp, body in cards:
    c.sticky(x, 40, 280, 250, color="green", weight="thin")
    c.emoji(cp, x + 140 - 32, 64, 64)
    c.text(x + 140, 166, title, scale="heading")
    c.text(x + 140, 206, body, scale="body")
c.save("02-three-rules.svg")

"""Excel でマクロを動かす 3 つの入口。VBE の F5、「開発」→「マクロ」の一覧、シートのボタン。どれも同じ Hello が動く（SOURCE-PATH-GOAL の合流）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1000, 420)
ways = [
    (40, "VBE で F5", "Sub の中にカーソルを置く\n（ツールバーの ▶ でも同じ）"),
    (150, "「開発」→「マクロ」", "一覧から Hello を選んで\n「実行」"),
    (260, "シートのボタン", "完成ブックの「あいさつする」\n（置き方は 5-3）"),
]
srcs = []
for y, title, note in ways:
    srcs.append(c.sticky(30, y, 340, 100, color="blue", weight="thin"))
    c.text(200, y + 38, title, scale="label", fill="#0f172a")
    c.text(200, y + 64, note, scale="caption")
hello = c.hexagon(500, 150, 160, 100, color="purple")
c.text(580, 208, "Hello", scale="heading", font="technical")
msg = c.node(860, 200, "「こんにちは、VBA！」", emoji_cp="1f4ac", w=100, h=80, label_scale="label")   # 💬
c.link((srcs[0], "e"), (hello, "n"), route="elbow")
c.link((srcs[1], "e"), (hello, "w"))
c.link((srcs[2], "e"), (hello, "s"), route="elbow")
c.link(hello, msg)
c.save("02-run-three-ways.svg")

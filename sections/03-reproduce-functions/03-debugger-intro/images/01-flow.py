"""3-3: デバッガーの流れ。置く → 実行する（黄色い行で止まる） → 変数を見る → 1 行進める（くり返す）→ 止める（SOURCE-PATH-GOAL + CYCLE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1080, 420)
c.text(540, 44, "デバッガーの使い方：5 つの手順", scale="heading")

STEPS = [
    ("red", "1f534", "① 置く", "止めたい行に\nブレークポイント"),        # 🔴
    ("yellow", "25b6", "② 実行する", "黄色い行で止まる\n黄色 ＝ これから\n実行する行"),  # ▶️
    ("blue", "1f50d", "③ 変数を見る", "i・total などの\nいまの値"),        # 🔍
    ("green", "23ed", "④ 1 行進める", "値がどう変わるか\nを確かめる"),     # ⏭️
    ("gray", "23f9", "⑤ 止める", "実行をやめて\n編集に戻る"),              # ⏹️
]
W, H, GAP, X0, Y = 168, 200, 46, 24, 80
cards = []
for k, (color, cp, title, body) in enumerate(STEPS):
    x = X0 + k * (W + GAP)
    cards.append(c.sticky(x, Y, W, H, color=color, weight="thin"))
    c.emoji(cp, x + W / 2 - 20, Y + 18, 40)
    c.text(x + W / 2, Y + 90, title, scale="label")
    c.text(x + W / 2, Y + 126, body, scale="caption")
for a, b in zip(cards, cards[1:]):
    c.link(a, b)
c.link((cards[3], "s"), (cards[2], "s"), route="elbow", label="何回もくり返す",
       label_scale="body", label_pos="below")
c.save("01-flow.svg")

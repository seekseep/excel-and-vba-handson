"""3-4: 原因の探し方の 4 手順。ずれを確かめる → 止める → 比べながら進める → 直して確かめる（合わなければ最初から）（CYCLE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1000, 440)
c.text(500, 44, "原因の探し方：4 つの手順をくり返す", scale="heading")

W, H = 330, 120
STEPS = [
    (90, 80, "yellow", "① ずれを確かめる", "期待（関数 68,580）と\n結果（VBA 67,330）の差は 1,250"),
    (580, 80, "blue", "② 怪しいところで止める", "合計を足している行に\nブレークポイント"),
    (580, 280, "green", "③ 比べながら進める", "変数を見て 1 行ずつ。\n正しく動くときとどこから違うか"),
    (90, 280, "purple", "④ 直して確かめる", "直したらもう一度実行して、\n期待どおりになるか見る"),
]
cards = []
for x, y, color, title, body in STEPS:
    cards.append(c.sticky(x, y, W, H, color=color, weight="thin"))
    c.text(x + W / 2, y + 40, title, scale="label")
    c.text(x + W / 2, y + 74, body, scale="caption")
c.link(cards[0], cards[1])
c.link(cards[1], cards[2])
c.link(cards[2], cards[3])
c.link(cards[3], cards[0], primary=False, dash="dashed")
c.text(280, 246, "合わなければ ① から", scale="caption", align="left")
c.save("02-cycle.svg")

"""5-3: Excel でシートにボタンを置く手順（SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(986, 240)
W, G = 200, 46
# (小さい字の行, 大きい字の行)
steps = [(None, "「開発」タブ"), (None, "「挿入」"),
         ("フォーム\nコントロールの", "「ボタン」"), ("空いているところを", "ドラッグ")]
prev = None
for k, (small, big) in enumerate(steps):
    x = 24 + k * (W + G)
    n = c.sticky(x, 60, W, 116, color="blue" if k < 3 else "yellow", rx=8, weight="thin")
    c.text(x + 24, 48, f"{k + 1}", scale="heading")
    if small:
        lines = small.count("\n") + 1
        c.text(x + W / 2, 162 - 36 - lines * 25 + 14, small, scale="caption")
        c.text(x + W / 2, 158, big, scale="label")
    else:
        c.text(x + W / 2, 126, big, scale="label")
    if prev:
        c.link(prev, n)
    prev = n
c.text(493, 218, "置くと「マクロの登録」が開く", scale="caption")
c.save("02-insert-button.svg")

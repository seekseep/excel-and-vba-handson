"""5-3: Excel のボタンとスプレッドシートのメニューは、同じ「イベント → 登録 → 手順のかたまり」（SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(960, 500)
W, G, X0 = 260, 60, 30
cols = [X0 + k * (W + G) for k in range(3)]
for x, head in zip(cols, ("イベント", "登録", "手順のかたまり")):
    c.text(x + W / 2, 44, head, scale="heading")

def box(x, y, color, lines, mono=()):
    n = c.sticky(x, y, W, 84, color=color, rx=8, weight="thin")
    n_lines = len(lines)
    for i, s in enumerate(lines):
        yy = y + 42 + (i - (n_lines - 1) / 2) * 28 + 6
        c.text(x + W / 2, yy, s, scale="caption", font="technical" if i in mono else None)
    return n

c.text(X0, 100, "Excel", scale="label", align="left")
e = [box(cols[0], 112, "yellow", ["シートのボタンを押す"]),
     box(cols[1], 112, "blue", ["「マクロの登録」で", "マクロを選ぶ"]),
     box(cols[2], 112, "green", ["Sub CreateInvoices()"], mono=(0,))]
c.text(X0, 256, "スプレッドシート", scale="label", align="left")
s = [box(cols[0], 268, "yellow", ["メニュー「請求書」の", "項目を押す"]),
     box(cols[1], 268, "blue", [".addItem('項目名',", "'関数名')"], mono=(0, 1)),
     box(cols[2], 268, "green", ["function", "createInvoices()"], mono=(0, 1))]
for row in (e, s):
    c.link(row[0], row[1])
    c.link(row[1], row[2])

op = c.sticky(cols[1], 400, W, 70, color="gray", rx=8, weight="thin", dash="dashed")
c.text(cols[1] + W / 2, 430, "開いたとき（onOpen）に", scale="caption")
c.text(cols[1] + W / 2, 456, "メニューを作って登録する", scale="caption")
c.link(op, s[1], primary=False)
c.save("01-button-vs-menu.svg")

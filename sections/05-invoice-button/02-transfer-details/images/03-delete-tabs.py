"""5-2: DeleteInvoices が消すシートと残すシート。「請求書_」で始まるかで分ける（SPLITTING）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(900, 360)
c.text(450, 42, "シート名が「請求書_」で始まるか", scale="label")

def tab(x, y, w, name, color):
    c.sticky(x, y, w, 44, color=color, rx=6, weight="thin")
    c.text(x + w / 2, y + 29, name, scale="caption")

c.text(40, 104, "◯ 残す", scale="heading", align="left", fill="#15803d")
x = 170
for name, w in (("売上", 80), ("顧客", 80), ("請求書テンプレート", 200)):
    tab(x, 76, w, name, "green")
    x += w + 12
c.text(x + 10, 104, "4 文字目が「テ」なので残る", scale="caption", align="left")

c.text(40, 194, "✕ 消す", scale="heading", align="left", fill="#b91c1c")
rows = [[("請求書_ひまわり商事", 210), ("請求書_あおば設計事務所", 250), ("請求書_みなと物産", 200)],
        [("請求書_さくら学習塾", 210), ("請求書_こもれび珈琲", 210)]]
for r, row in enumerate(rows):
    x = 170
    for name, w in row:
        tab(x, 166 + r * 60, w, name, "red")
        x += w + 12
c.text(450, 330, "VBA: Left(ws.Name, 4) = \"請求書_\"　GAS: startsWith('請求書_')", scale="caption", font="technical")
c.save("03-delete-tabs.svg")

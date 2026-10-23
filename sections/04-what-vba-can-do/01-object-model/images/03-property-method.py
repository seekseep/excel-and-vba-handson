"""4-1: オブジェクトの 2 種類の要素。プロパティ（状態）とメソッド（動作）を左右に分ける（SPLITTING）。

キャンバスを表示幅（720px）にして、座標と字の大きさを表示どおりにしている。
"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(720, 300)
c.text(360, 34, "部品（オブジェクト）の後ろに . を付けて書くもの", scale="label")

groups = [
    (16, "blue", "プロパティ ＝ 部品の状態", "読んだり書き換えたりできる",
     [(".Value", "値"), (".Name", "名前"), (".Interior.Color", "背景色")]),
    (368, "purple", "メソッド ＝ 部品にさせる動作", "部品に命令して動かす",
     [(".Copy", "コピー"), (".Delete", "削除"), (".Add", "追加")]),
]
for x, color, head, sub, rows in groups:
    w = 336
    c.sticky(x, 56, w, 228, color=color, weight="thin")
    c.text(x + w / 2, 90, head, scale="label", fill="#0f172a")
    c.text(x + w / 2, 116, sub, scale="body")
    for i, (code, meaning) in enumerate(rows):
        y = 134 + i * 48
        c.sticky(x + 24, y, w - 48, 38, color=color, rx=8, weight="thin")
        c.text(x + 40, y + 25, code, scale="body", font="technical", weight="700",
               fill="#0f172a", align="left")
        c.text(x + w - 40, y + 25, meaning, scale="body", align="right")
c.save("03-property-method.svg")

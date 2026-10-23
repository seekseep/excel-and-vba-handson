"""4-2: SplitByCustomer のデータの動き。外側のくり返しで顧客を 1 社ずつ取り出し、
内側のくり返しで売上の行を上から見て、同じ顧客の行だけを顧客のシートの outRow 行目に写す
（SOURCE-PATH-GOAL + SPLITTING）。

キャンバスを表示幅（720px）にして、表の字（figtable）も表示どおりの大きさにしている。
"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(720, 720)
H = 30
hit = "yellow"

# ① 外側のくり返し: H5:H9 の顧客名
c.text(16, 36, "① 外側のくり返し c = 5 → 9", scale="label", align="left")
c.text(16, 62, "H 列から顧客を 1 社ずつ", scale="body", align="left")
cy = 80
cust = table(c, 16, cy, [
    ["行", "H"],
    ["4", "顧客"],
    ["5", "ひまわり商事"],
    ["6", "あおば設計事務所"],
    ["7", "みなと物産"],
    ["8", "さくら学習塾"],
    ["9", "こもれび珈琲"],
], [32, 140], color="gray", h=H, scale="body",
    fills={(1, 0): "gray", (1, 1): "gray", (2, 0): hit, (2, 1): hit})
c.text(16, cy + H * 7 + 34, "c = 5 のとき\ncustomerName は\n「ひまわり商事」", scale="body", align="left")

# ② 内側のくり返し: 売上の行を上から
sx, sy = 300, 80
c.text(sx, 36, "② 内側のくり返し i = 2 → 31", scale="label", align="left")
c.text(sx, 62, "B 列が customerName と同じ行を探す", scale="body", align="left")
sales = table(c, sx, sy, [
    ["行", "A", "B", "C〜F"],
    ["1", "日付", "顧客", "…"],
    ["2", "10/02", "あおば設計事務所", "…"],
    ["3", "10/06", "みなと物産", "…"],
    ["4", "10/07", "ひまわり商事", "…"],
    ["5", "10/07", "ひまわり商事", "…"],
    ["6", "10/07", "みなと物産", "…"],
    ["7", "10/08", "さくら学習塾", "…"],
    ["8", "10/08", "ひまわり商事", "…"],
    ["⋮", "⋮", "⋮", "⋮"],
    ["31", "10/29", "みなと物産", "…"],
], [36, 64, 150, 56], color="orange", h=H, scale="body",
    fills={**{(1, k): "gray" for k in range(4)},
           **{(r, k): hit for r in (4, 5, 8) for k in range(4)}})
c.link((cust.right + 6, cy + H * 2.5), (sx - 6, sy + H * 4.5), label="同じ名前", label_pos="above")
c.link((sales.right + 20, sy + H * 2), (sales.right + 20, sy + H * 11), label="i", primary=False)

# ③ 写し先: 顧客のシート
ox, oy = 16, 520
c.text(ox, oy - 52, "③ 同じ行だけ A〜F をまるごとコピー", scale="label", align="left")
c.text(ox, oy - 24, "足したシート「ひまわり商事」に書き写す", scale="body", align="left")
out = table(c, ox, oy, [
    ["行", "A", "B", "C〜F", "どこから"],
    ["1", "日付", "顧客", "…", "売上の見出し"],
    ["2", "10/07", "ひまわり商事", "…", "売上の 4 行目"],
    ["3", "10/07", "ひまわり商事", "…", "売上の 5 行目"],
    ["4", "10/08", "ひまわり商事", "…", "売上の 8 行目"],
], [36, 64, 150, 56, 120], color="yellow", h=H, scale="body",
    fills={**{(1, k): "gray" for k in range(5)}})
c.link((400, sales.bottom + 6), (400, oy - 6))

nx = out.right + 24
c.text(nx, oy + H * 2 + 20, "outRow：書き写す先の行", scale="body", align="left", fill="#0f172a")
c.text(nx, oy + H * 3 + 18, "2 から始めて、1 行書くたびに\n1 増やす（2 → 3 → 4）", scale="body", align="left")
c.text(360, 700, "c = 6（あおば設計事務所）でも同じことをくり返す。outRow は 2 から数え直す",
       scale="body")
c.save("01-split-flow.svg")

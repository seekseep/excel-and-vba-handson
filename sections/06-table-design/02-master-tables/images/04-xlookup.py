"""6-2: XLOOKUP で売上の顧客名から顧客シートの住所を引く。3 つの引数を列の色に合わせる（LINK）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from genfig.text import text_extent
from figtable import table

c = Canvas(900, 600)

# 式を引数ごとのチップに分けて、色で表の列と対応させる
parts = [("=XLOOKUP(", None, None), ("B2", "yellow", "探す値"), (",", None, None),
         ("顧客!A:A", "blue", "探す列"), (",", None, None), ("顧客!C:C", "green", "返す列"), (")", None, None)]
pad, gap = 14, 6
widths = [text_extent(s, "label", font="technical")[0] + (2 * pad if col else 0) for s, col, _ in parts]
x = 450 - (sum(widths) + gap * (len(parts) - 1)) / 2
for (s, col, note), w in zip(parts, widths):
    if col:
        c.sticky(x, 24, w, 46, color=col, rx=6, weight="thin")
        c.text(x + w / 2, 96, note, scale="caption")
    c.text(x + w / 2, 55, s, scale="label", font="technical")
    x += w + gap

h = 38
c.text(100, 175, "売上", scale="label", align="right")
sx = 120
table(c, sx, 150, [
    ["A 日付", "B 顧客", "C 商品", "D 住所（この式）"],
    ["10/07", "ひまわり商事", "ホッチキス", "東京都中央区ひまわり…"],
], [70, 130, 90, 210], color="gray", h=h, scale="body",
    fills={(0, k): "gray" for k in range(4)} | {(1, 1): "yellow", (1, 3): "green"})

c.text(100, 345, "顧客", scale="label", align="right")
ky = 320
table(c, sx, ky, [
    ["A 顧客名", "B 郵便番号", "C 住所"],
    ["ひまわり商事", "〒104-0000", "東京都中央区ひまわり…"],
    ["あおば設計事務所", "〒160-0000", "東京都新宿区あおば台…"],
    ["みなと物産", "〒231-0000", "神奈川県横浜市中区…"],
], [140, 110, 210], color="gray", h=h, scale="body",
    fills={(0, 0): "blue", (0, 1): "gray", (0, 2): "green", (1, 0): "yellow", (1, 2): "green"})

b2 = (sx + 70 + 65, 150 + h * 2)
c.link(b2, (sx + 70, ky), label="① A 列から探す", label_pos="above")
# ② 見つかった行を枠で囲む
c.raw(f'<rect x="{sx - 4}" y="{ky + h - 4}" width="{460 + 8}" height="{h + 8}" rx="6" fill="none" '
      f'stroke="#f59e0b" stroke-width="3" stroke-dasharray="7 6"/>')
rx = sx + 560
c.link((sx + 464, ky + h * 1.5), (rx, ky + h * 1.5), head=False)
c.link((rx, ky + h * 1.5), (rx, 150 + h * 1.5), head=False)
c.link((rx, 150 + h * 1.5), (sx + 500, 150 + h * 1.5))
c.text(rx + 16, 262, "② 同じ行の\n　 C 列を\n③ 返す", scale="label", align="left")

note = c.sticky(120, 510, 660, 60, color="yellow", weight="thin", rx=8)
c.text(450, 548, "第 5 章のマクロも同じつなぎ方（顧客名で 2 枚のシートをつなぐ）", scale="caption")
c.save("04-xlookup.svg")

"""5-2: 売上の列と請求書の列の対応、売上の行 i と書き写し先の行 lineRow の対応（LINK + SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(900, 700)
RED = "#b91c1c"

# 上: 列の対応
c.text(40, 44, "列の対応", scale="heading", align="left")
x0, cw, h = 150, 110, 34
c.text(x0 - 16, 94, "売上", scale="label", align="right")
table(c, x0, 64, [
    ["A", "B", "C", "D", "E", "F"],
    ["日付", "顧客", "商品", "単価", "数量", "金額"],
], [cw] * 6, color="blue", h=h, scale="body", fills={(1, 1): "red"})
c.text(x0 - 16, 254, "請求書", scale="label", align="right")
table(c, x0, 224, [
    ["A", "B", "C", "D", "E"],
    ["日付", "商品", "単価", "数量", "金額"],
], [cw] * 5, color="orange", h=h, scale="body")
top, bot = 64 + h * 2, 224
c.link((x0 + cw * 0.5, top), (x0 + cw * 0.5, bot))
for k in range(2, 6):
    c.link((x0 + cw * (k + 0.5), top), (x0 + cw * (k - 0.5), bot))
c.text(x0 + cw * 1.5, top + 34, "✕ 書かない", scale="caption", fill=RED)

# 下: 行の対応
c.text(40, 370, "行の対応（ひまわり商事のとき）", scale="heading", align="left")
ly, rh = 394, 34
table(c, 60, ly, [
    ["i", "売上の顧客（B 列）"],
    ["2", "あおば設計事務所"],
    ["3", "みなと物産"],
    ["4", "ひまわり商事"],
    ["5", "ひまわり商事"],
    ["6", "みなと物産"],
    ["7", "さくら学習塾"],
    ["8", "ひまわり商事"],
], [50, 170], color="blue", h=rh, scale="body",
    fills={(r, k): "yellow" for r in (3, 4, 7) for k in (0, 1)})
ry = ly + rh * 2
table(c, 590, ry, [
    ["lineRow", "請求書の商品（B 列）"],
    ["12", "ホッチキス"],
    ["13", "コピー用紙"],
    ["14", "ノート"],
], [90, 190], color="orange", h=rh, scale="body")
for src, dst in ((3, 1), (4, 2), (7, 3)):
    c.link((280, ly + rh * (src + 0.5)), (590, ry + rh * (dst + 0.5)), primary=False)
c.text(730, ry + rh * 4 + 40, "一致した行だけ書き写し、\nそのたびに lineRow を 1 つ進める", scale="caption")
c.save("02-column-map.svg")

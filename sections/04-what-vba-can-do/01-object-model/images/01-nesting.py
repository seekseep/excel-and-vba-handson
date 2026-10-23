"""4-1: VBA から見た Excel の入れ子と、それを `.` でたどる書き方（CONTAINER + PATH）。

上の段のコードの部品と、下の段の箱の左端をそろえ、「外側から内側へ `.` でたどる」ことを見せる。
キャンバスを表示幅（720px）にして、座標と字の大きさを表示どおりにしている。
"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas
from genfig.text import text_width

c = Canvas(720, 500)

parts = [
    ("ThisWorkbook", "このブックの", "green"),
    ('Worksheets("売上")', "売上シートの", "blue"),
    ('Range("B2")', "B2 セルの", "yellow"),
    ("Value", "値", "orange"),
]
dot_w, pad, chip_y, chip_h = 20, 10, 56, 40

# 札の位置を先に決める（省略できる範囲の枠を札より先に描くため）
x, chips = 56, []
for i, (code, _, _) in enumerate(parts):
    if i:
        x += dot_w
    w = text_width(code, "body", font="technical", weight="700") + pad * 2
    chips.append((x, w))
    x += w

# 先頭の 2 つは省略できる
omit_l = chips[0][0] - 8
omit_r = chips[1][0] + chips[1][1] + dot_w - 4
c.sticky(omit_l, chip_y - 10, omit_r - omit_l, chip_h + 20, color="gray", rx=12,
         weight="thin", dash="dashed")
c.text((omit_l + omit_r) / 2, chip_y - 22, "ここを省略すると「いま開いているシート」になる", scale="body")

for i, ((code, reading, color), (cx, w)) in enumerate(zip(parts, chips)):
    if i:
        c.text(cx - dot_w / 2, chip_y + 28, ".", scale="heading", font="technical", fill="#0f172a")
    c.sticky(cx, chip_y, w, chip_h, color=color, rx=8, weight="thin")
    c.text(cx + w / 2, chip_y + 25, code, scale="body", font="technical", weight="700", fill="#0f172a")
    c.text(cx + w / 2, chip_y + chip_h + 34, reading, scale="body")

# 入れ子の箱。各箱の左端を、上の札の左端にそろえる
top, right, bottom, step, shrink = 156, 704, 484, 42, 14
c.sticky(16, top, right - 16, bottom - top, color="gray", weight="thin")
c.text(32, top + 28, "Excel そのもの（Application）", scale="label", align="left")
boxes = [
    ("ブック（Workbook）", "このマクロが\n入っているブック"),
    ("シート（Worksheet）", "「売上」という\n名前のシート"),
    ("セル・範囲（Range）", "Cells(2, 2) とも書ける"),
]
for k, ((cx, _), (head, note)) in enumerate(zip(chips, boxes), start=1):
    bx, by = cx - 12, top + step * k
    c.sticky(bx, by, right - shrink * k - bx, bottom - shrink * k - by,
             color=parts[k - 1][2], weight="thin")
    c.text(bx + 14, by + 28, head, scale="label", align="left")
    c.text(bx + 14, by + (58 if k == 3 else 76), note, scale="body", align="left")

# 値はセルの中の 1 項目
vx, vw = chips[3]
tag_w = 100
c.sticky(vx - 4, top + step * 3 + 92, tag_w, 40, color="orange", rx=8, weight="thin")
c.text(vx - 4 + tag_w / 2, top + step * 3 + 117, "値（Value）", scale="body", fill="#0f172a")
c.save("01-nesting.svg")

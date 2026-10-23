"""4-1: Apps Script から見たスプレッドシートの入れ子と、それを `.` でたどる書き方（CONTAINER + PATH）。

01-nesting（VBA）と同じ形。上の段のコードの部品と、下の段の箱の左端をそろえる。
キャンバスを表示幅（720px）にして、座標と字の大きさを表示どおりにしている。
"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas
from genfig.text import text_width

c = Canvas(720, 500)

parts = [
    ("ss", "このスプレッドシートの", "green"),
    ("getSheetByName('売上')", "売上シートの", "blue"),
    ("getRange('B2')", "B2 セルの", "yellow"),
    ("getValue()", "値を読む", "orange"),
]
dot_w, pad, chip_y, chip_h = 20, 10, 56, 40

x, chips = 66, []
for i, (code, _, _) in enumerate(parts):
    if i:
        x += dot_w
    w = max(80, text_width(code, "body", font="technical", weight="700") + pad * 2)
    chips.append((x, w))
    x += w

c.text(chips[0][0], chip_y - 22, "ss は SpreadsheetApp.getActiveSpreadsheet()（デモの 1 行目で作る）",
       scale="body", align="left")

for i, ((code, reading, color), (cx, w)) in enumerate(zip(parts, chips)):
    if i:
        c.text(cx - dot_w / 2, chip_y + 28, ".", scale="heading", font="technical", fill="#0f172a")
    c.sticky(cx, chip_y, w, chip_h, color=color, rx=8, weight="thin")
    c.text(cx + w / 2, chip_y + 25, code, scale="body", font="technical", weight="700", fill="#0f172a")
    c.text(cx + w / 2, chip_y + chip_h + 34, reading, scale="body")

top, right, bottom, step, shrink = 156, 704, 484, 42, 14
c.sticky(16, top, right - 16, bottom - top, color="gray", weight="thin")
c.text(32, top + 28, "Apps Script から見た Google スプレッドシート（SpreadsheetApp）", scale="label", align="left")
boxes = [
    ("スプレッドシート（Spreadsheet）", None),
    ("シート（Sheet）", "「売上」という\n名前のシート"),
    ("セル・範囲（Range）", "getRange(2, 2) とも書ける"),
]
for k, ((cx, _), (head, note)) in enumerate(zip(chips, boxes), start=1):
    bx, by = cx - 12, top + step * k
    c.sticky(bx, by, right - shrink * k - bx, bottom - shrink * k - by,
             color=parts[k - 1][2], weight="thin")
    c.text(bx + 14, by + 28, head, scale="label", align="left")
    if note:
        c.text(bx + 14, by + (58 if k == 3 else 76), note, scale="body", align="left")

vx, vw = chips[3]
tag_w = 76
c.sticky(vx - 4, top + step * 3 + 92, tag_w, 40, color="orange", rx=8, weight="thin")
c.text(vx - 4 + tag_w / 2, top + step * 3 + 117, "値", scale="body", fill="#0f172a")
c.save("02-nesting-gas.svg")

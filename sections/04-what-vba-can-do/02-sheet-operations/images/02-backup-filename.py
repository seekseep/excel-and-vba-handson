"""4-2: SaveBackup で作るファイル名の内訳。5 つの部品を & でつないで 1 本のパスにする（PART-WHOLE）。

上の段が出来上がるパス、下の段がそれぞれの部品を作るコード（コード → 値の向きに矢印）。
キャンバスを表示幅（720px）にして、座標と字の大きさを表示どおりにしている。
"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas
from genfig.text import text_width

c = Canvas(720, 370)
c.text(360, 34, "5 つの部品を & でつないで、1 本のパスにする（例は Windows）", scale="label")

segs = [
    ("C:\\Users\\…\\Documents", "このブックのフォルダ", "ThisWorkbook.Path", None, 0, "green"),
    ("\\", "区切り", "Application.PathSeparator", "Windows は \\、Mac は /", 2, "gray"),
    ("売上_バックアップ_", "決まった文字", '"売上_バックアップ_"', None, 0, "blue"),
    ("20261022_143005", "今の日時", 'Format(Now, "yyyymmdd_hhnnss")', "年月日_時分秒（nn は分）", 1, "yellow"),
    (".xlsm", "拡張子", '".xlsm"', None, 0, "purple"),
]
pad, bar_y, bar_h = 8, 96, 44
widths = [text_width(v, "body", font="technical", weight="700") + pad * 2 for v, *_ in segs]
x = (720 - sum(widths)) / 2
code_y = [bar_y + bar_h + 56, bar_y + bar_h + 116, bar_y + bar_h + 176]
for (value, role, code, note, level, color), w in zip(segs, widths):
    cx = x + w / 2
    c.sticky(x, bar_y, w, bar_h, color=color, rx=4, weight="thin")
    c.text(cx, bar_y + 27, value, scale="body", font="technical", weight="700", fill="#0f172a")
    c.text(cx, bar_y - 14, role, scale="body")
    y = code_y[level]
    c.link((cx, y - 22), (cx, bar_y + bar_h + 4), primary=False, dash="dotted")
    c.text(cx, y, code, scale="body", font="technical", fill="#0f172a")
    if note:
        c.text(cx, y + 24, note, scale="body")
    x += w
c.save("02-backup-filename.svg")

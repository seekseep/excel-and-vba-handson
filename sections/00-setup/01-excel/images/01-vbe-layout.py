"""VBE の画面構成。Excel から Alt + F11 で別のウィンドウとして開き、左のプロジェクトの Module1 に書く（CONTAINER / PART-WHOLE）。

0-1 と 2-1 に同じ図を置く。直すときは両方の images/ で走らせる。
"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1040, 470)

# 左: Excel のウィンドウ
excel = c.sticky(20, 120, 200, 200, color="green", weight="thin", rx=8)
c.text(120, 152, "Excel", scale="label")
for r in range(4):
    for k in range(3):
        c.sticky(38 + k * 56, 172 + r * 32, 52, 28, color="gray", weight="thin", rx=3)
c.text(120, 352, "いつもの画面", scale="caption")

# 右: VBE のウィンドウ
vbe = c.sticky(330, 30, 680, 410, color="gray", weight="thin", rx=8)
c.text(670, 64, "VBE（別のウィンドウで開く）", scale="label")
c.text(350, 102, "メニュー「挿入」→「標準モジュール」で Module1 ができる", scale="caption", align="left")

# プロジェクト欄
c.sticky(350, 122, 230, 298, color="blue", weight="thin", rx=6)
c.text(465, 152, "プロジェクト", scale="label")
c.text(368, 192, "VBAProject（Book1）", scale="caption", align="left")
c.text(388, 226, "Sheet1", scale="caption", align="left")
c.text(388, 258, "ThisWorkbook", scale="caption", align="left")
c.text(388, 292, "標準モジュール", scale="caption", align="left")
mod = c.sticky(398, 306, 160, 36, color="yellow", rx=6)
c.text(478, 330, "Module1", scale="label")
c.text(465, 396, "Sheet1 などは別の用途", scale="caption")

# コード欄
code = c.sticky(610, 122, 380, 298, color="gray", weight="thin", rx=6)
c.raw('<rect x="616" y="128" width="368" height="286" rx="4" fill="#ffffff"/>')
c.text(800, 152, "コード（Module1）", scale="label")
c.text(634, 200, "Sub Hello()", scale="body", align="left", font="technical")
c.text(666, 240, "MsgBox \"こんにちは\"", scale="body", align="left", font="technical")
c.text(634, 280, "End Sub", scale="body", align="left", font="technical")
c.text(800, 370, "ここに書く", scale="heading", fill="#b45309")

c.link((excel, "e"), (330, 220), label="Alt + F11", label_pos="above", label_scale="body")
c.text(275, 262, "Mac は\nOption + F11", scale="caption")
c.link((mod, "e"), (610, 324), primary=False)
c.save("01-vbe-layout.svg")

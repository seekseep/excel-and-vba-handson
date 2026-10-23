"""スプレッドシートと Apps Script エディタは別のタブ。拡張機能 → Apps Script で開き、エディタで実行した結果はスプレッドシートのタブに出る（LINK + CYCLE）。

0-2 と 2-1 に同じ図を置く。直すときは両方の images/ で走らせる。
"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1040, 450)


def tab(x, title, color):
    c.sticky(x + 16, 30, 220, 44, color=color, weight="thin", rx=8)
    c.text(x + 126, 60, title, scale="body", fill="#0f172a")
    return c.sticky(x, 70, 400, 340, color="gray", weight="thin", rx=8)


sheet = tab(30, "スプレッドシート", "green")
c.sticky(46, 90, 368, 40, color="green", weight="thin", rx=4)
c.text(60, 117, "ファイル  編集  …  拡張機能", scale="caption", align="left")
menu = c.sticky(240, 136, 170, 44, color="yellow", rx=6)
c.text(325, 164, "Apps Script", scale="body", fill="#0f172a")
for r in range(4):
    for k in range(4):
        c.sticky(52 + k * 90, 200 + r * 34, 84, 30, color="gray", weight="thin", rx=2)
msg = c.sticky(110, 250, 240, 64, color="blue", rx=8)
c.text(230, 290, "こんにちは", scale="body", fill="#0f172a")
c.text(230, 388, "結果（小窓・セルの値）はここに出る", scale="caption")

editor = tab(610, "Apps Script", "blue")
c.sticky(626, 90, 368, 40, color="blue", weight="thin", rx=4)
run = c.sticky(800, 94, 90, 32, color="yellow", rx=6)
c.text(845, 116, "▶ 実行", scale="caption", fill="#0f172a")
c.text(640, 117, "hello ▾", scale="caption", align="left")
c.text(634, 170, "コード.gs", scale="body", align="left", fill="#0f172a")
c.text(634, 214, "function hello() {", scale="body", align="left", font="technical")
c.text(660, 248, "SpreadsheetApp.getUi()", scale="body", align="left", font="technical")
c.text(660, 280, "  .alert('こんにちは');", scale="body", align="left", font="technical")
c.text(634, 314, "}", scale="body", align="left", font="technical")
c.text(810, 388, "コードを書いて、関数を選んで実行", scale="caption")

c.link((menu, "e"), (610, 158))
c.text(520, 210, "① 拡張機能 →\nApps Script で\n新しいタブに開く", scale="caption")
c.link((610, 282), (msg, "e"))
c.text(520, 330, "② ▶ 実行の結果は\nスプレッドシートの\nタブに戻って見る", scale="caption")
c.save("02-two-tabs.svg")

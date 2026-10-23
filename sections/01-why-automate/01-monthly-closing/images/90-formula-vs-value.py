"""関数のセルは式が残るので元のデータを直すと計算し直す。VBA が書いたセルは値だけなので古いまま（BALANCE / SOURCE-PATH-GOAL）。

3-1・3-2・3-4 と 1-1 で同じ図を使う。直したらこのファイルと SVG を 4 か所にコピーする。
"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(1000, 490)
c.text(500, 44, "関数は「式」が残り、VBA は「値」だけが残る", scale="heading")

W = [110, 110, 130]
HEAD = ["D2 単価", "E2 数量", "F2 金額"]

for x0, color, title, inside, after, mark, note in [
    (60, "yellow", "関数で出したセル", "F2 の中身は式 =D2*E2", "1500", "green",
     "◯ すぐに計算し直す"),
    (590, "blue", "VBA が書いたセル", "F2 の中身は値 1250", "1250", "red",
     "✕ 古いまま\nもう一度実行するまで変わらない"),
]:
    cx = x0 + sum(W) / 2
    c.text(cx, 100, title, scale="label")
    before = table(c, x0, 120, [HEAD, ["250", "5", "1250"]], W, color=color, h=36,
                   fills={(1, 2): color}, scale="body")
    c.text(cx, 232, inside, scale="label")
    after_t = table(c, x0, 320, [HEAD, ["300", "5", after]], W, color=color, h=36,
                    fills={(1, 0): "orange", (1, 2): mark}, scale="body")
    c.link((cx, 250), after_t, label="D2 を 300 に書き換える", label_scale="body")
    c.text(cx, 425, note, scale="label", fill="#15803d" if mark == "green" else "#b91c1c")
c.save("90-formula-vs-value.svg")

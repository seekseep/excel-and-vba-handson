"""代入。右側を先に計算して、その結果を左の箱（変数）に入れ直す。total は 0 → 1250 → 3750（CONTAINER / CYCLE）。

3-2 と 2-2 で同じ図を使う。直したらこのファイルと SVG を 2 か所にコピーする。
"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1000, 470)
c.text(500, 44, "「=」は「右側を計算して、左の箱に入れる」", scale="heading")

left = c.sticky(250, 110, 130, 60, color="blue", rx=10, weight="thin")
c.text(315, 148, "total", scale="title", font="technical")
c.text(420, 150, "=", scale="title")
right = c.sticky(460, 110, 290, 60, color="yellow", rx=10, weight="thin")
c.text(605, 148, "total + 金額", scale="title", font="technical")
c.text(605, 96, "① 先に右側を計算する", scale="label")
c.link((right, "s"), (left, "s"), route="elbow", label="② 結果を左の total に入れ直す",
       label_scale="body", label_pos="below")

c.text(500, 300, "くり返すたびに、箱の中身が入れ替わる", scale="label")
boxes = []
for k, (x, v, cap) in enumerate([(90, "0", "はじめ"), (430, "1250", "2 行目を足した"),
                                 (770, "3750", "3 行目を足した")]):
    b = c.sticky(x, 330, 140, 80, color="blue", rx=10, weight="thin")
    c.text(x + 70, 352, "total", scale="caption", font="technical")
    c.text(x + 70, 398, v, scale="heading")
    c.text(x + 70, 440, cap, scale="body")
    boxes.append(b)
c.link(boxes[0], boxes[1], label="0 + 1250", label_scale="body", label_pos="above")
c.link(boxes[1], boxes[2], label="1250 + 2500", label_scale="body", label_pos="above")
c.save("92-assign.svg")

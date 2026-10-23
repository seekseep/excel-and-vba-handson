"""5-2: 顧客のくり返しの中に、売上のくり返しが入る（CONTAINER + CYCLE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(860, 660)
outer = c.sticky(40, 30, 780, 580, color="blue", weight="thin", rx=12)
c.text(64, 70, "For c = 2 To 最後の顧客の行", scale="heading", align="left", font="technical")
c.text(64, 102, "顧客を 1 社ずつ", scale="caption", align="left")

copy = c.sticky(80, 124, 520, 50, color="gray", weight="thin", rx=8)
c.text(100, 156, "テンプレートをコピーして宛名を書く（5-1）", scale="caption", align="left")

init = c.sticky(80, 192, 520, 60, color="yellow", weight="thin", rx=8)
c.text(100, 230, "lineRow = 12 ／ total = 0", scale="label", align="left", font="technical")
c.text(640, 216, "1 社ごとに\nここで初期化する", scale="caption", align="left", fill="#b45309")

inner = c.sticky(80, 272, 700, 210, color="green", weight="thin", rx=12)
c.text(104, 310, "For i = 2 To 最後の売上の行", scale="label", align="left", font="technical")
c.text(104, 338, "売上を 1 行ずつ", scale="caption", align="left")
q = c.node(250, 400, "この顧客の行？", shape="diamond", color="yellow", w=230, h=90, label_scale="caption")
w = c.sticky(450, 360, 300, 80, color="green", rx=8, weight="thin")
c.text(600, 392, "請求書の lineRow 行に書き写す", scale="caption")
c.text(600, 420, "total に足す・lineRow + 1", scale="caption")
c.link(q, w, label="はい")
c.text(740, 470, "Next i", scale="label", align="right", font="technical")

tot = c.sticky(80, 502, 520, 50, color="orange", weight="thin", rx=8)
c.text(100, 534, "E33 と B9 に total を書く", scale="caption", align="left")
c.text(796, 590, "Next c", scale="label", align="right", font="technical")
c.text(430, 642, "GAS では for (let c …) の中に for (let i …)。内側の } が先に閉じる", scale="caption")
c.save("01-nested-loop.svg")

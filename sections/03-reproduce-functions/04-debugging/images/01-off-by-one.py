"""3-4: 1 行ずれのバグ。正しい SumTotal は 2 行目から足して 68,580。バグありは 3 行目から足すので 2 行目の 1,250 が抜けて 67,330（PART-WHOLE / SCALE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas
from genfig.tokens import PALETTE

c = Canvas(1000, 430)
c.text(500, 44, "3 行目から足し始めると、2 行目の分だけ足りない", scale="heading")

CW, CH, X0 = 22, 34, 190


def strip(y, start, color):
    for k, row in enumerate(range(2, 32)):
        x = X0 + k * CW
        used = row >= start
        p = PALETTE[color if used else "red"]
        dash = '' if used else ' stroke-dasharray="4 3"'
        c.raw(f'<rect x="{x}" y="{y}" width="{CW}" height="{CH}" fill="{p["bg"] if used else "#ffffff"}" '
              f'stroke="{p["border"]}" stroke-width="1.5"{dash}/>')
        if row in (2, 3, 31):
            c.text(x + CW / 2, y + CH / 2 + 5, str(row), scale="caption")


c.text(X0 + 15 * CW, 88, "2〜31 行目の金額（1 マス ＝ 1 行）", scale="caption")
c.text(40, 120, "SumTotal", scale="label", align="left", font="technical")
c.text(40, 146, "i = 2 から", scale="caption", align="left")
strip(100, 2, "blue")
c.text(X0 + 30 * CW + 16, 123, "68,580", scale="heading", align="left")

c.text(40, 250, "SumTotal", scale="label", align="left", font="technical")
c.text(40, 276, "WithBug", scale="label", align="left", font="technical")
c.text(40, 300, "i = 3 から", scale="caption", align="left")
strip(230, 3, "orange")
c.text(X0 + 30 * CW + 16, 253, "67,330", scale="heading", align="left", fill="#b91c1c")
miss = c.sticky(X0 + 30, 300, 300, 46, color="red", rx=10, weight="thin")
c.text(X0 + 180, 329, "2 行目の 1,250 が足されない", scale="body")
c.link(miss, (X0 + CW / 2, 230 + CH + 4), primary=False)

c.text(500, 400, "差は 68,580 − 67,330 ＝ 1,250。これは 2 行目の金額と同じ", scale="label")
c.save("01-off-by-one.svg")

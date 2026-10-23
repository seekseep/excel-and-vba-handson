"""5-3: PDF ができる場所。ブック（スプレッドシート）と同じフォルダ（CONTAINER）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(960, 390)
RED = "#b91c1c"

c.text(230, 40, "Excel", scale="heading")
c.sticky(30, 60, 400, 200, color="yellow", weight="thin", rx=12)
c.text(50, 92, "ブックを保存したフォルダ（PC の中）", scale="caption", align="left")
c.node(130, 175, "invoice.xlsm", emoji_cp="1f4d7", w=70, h=60, label_scale="caption")   # 📗
c.node(320, 175, "請求書_〇〇.pdf × 5", emoji_cp="1f4c4", w=70, h=60, label_scale="caption")  # 📄
c.node(90, 330, "", emoji_cp="2601", w=60, h=50)   # ☁️
c.text(140, 322, "✕ OneDrive 上のブック", scale="label", align="left", fill=RED)
c.text(140, 352, "保存先が https://… になり、書き出せないことがある", scale="caption", align="left")

c.text(730, 40, "スプレッドシート", scale="heading")
c.sticky(530, 60, 400, 200, color="green", weight="thin", rx=12)
c.text(550, 92, "スプレッドシートのあるドライブのフォルダ", scale="caption", align="left")
c.node(630, 175, "スプレッドシート", emoji_cp="1f4ca", w=70, h=60, label_scale="caption")   # 📊
c.node(820, 175, "請求書_〇〇.pdf × 5", emoji_cp="1f4c4", w=70, h=60, label_scale="caption")
c.text(730, 322, "ふつうはマイドライブ", scale="caption")
c.save("03-pdf-location.svg")

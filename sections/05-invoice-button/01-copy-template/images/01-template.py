"""第5章: 請求書テンプレートのセル配置。どのセルに何を書くか（PART-WHOLE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(760, 640)
c.sticky(60, 30, 640, 580, color="gray", weight="thin", rx=8)
c.text(380, 82, "請 求 書", scale="title")

c.sticky(500, 104, 180, 40, color="yellow", rx=6, weight="thin")
c.text(590, 130, "E3  発行日", scale="caption")

c.sticky(90, 160, 300, 110, color="yellow", rx=6, weight="thin")
c.text(240, 198, "A5  宛名（〇〇 御中）", scale="body")
c.text(240, 226, "A6  住所", scale="caption")
c.text(240, 250, "A7  ご担当者", scale="caption")

c.sticky(420, 160, 260, 110, color="gray", rx=6, weight="thin")
c.text(550, 208, "そよかぜ文具店", scale="body")
c.text(550, 236, "（テンプレートのまま）", scale="caption")

c.sticky(90, 290, 300, 50, color="orange", rx=6, weight="thin")
c.text(240, 321, "B9  ご請求金額", scale="body")

c.sticky(90, 360, 590, 40, color="gray", rx=6, weight="thin")
c.text(385, 386, "11 行目  日付 ／ 商品 ／ 単価 ／ 数量 ／ 金額", scale="caption")
c.sticky(90, 404, 590, 130, color="blue", rx=6, weight="thin")
c.text(385, 462, "12〜31 行目  明細（この顧客の売上を書き写す）", scale="body")
c.text(385, 490, "A 日付  B 商品  C 単価  D 数量  E 金額", scale="caption")

c.sticky(500, 550, 180, 40, color="orange", rx=6, weight="thin")
c.text(590, 576, "E33  合計", scale="caption")
c.save("01-template.svg")

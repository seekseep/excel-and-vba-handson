"""6-2: 顧客名を毎回手で打つと表記がゆれ、集計で別の会社に分かれる（SPLITTING）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(960, 330)
c.text(40, 60, "売上（顧客名を手で入力）", scale="heading", align="left")
sales = table(c, 40, 90, [
    ["日付", "顧客", "金額"],
    ["10/07", "ひまわり商事", "1,360"],
    ["10/08", "㈱ひまわり商事", "360"],
    ["10/13", "ひまわり商事␣", "1,100"],
    ["10/15", "ひまわり商事", "6,800"],
    ["10/21", "ヒマワリ商事", "900"],
], [70, 150, 80], color="orange", h=34,
    fills={(2, 1): "red", (3, 1): "red", (5, 1): "red"})
c.text(600, 60, "顧客ごとの合計", scale="heading", align="left")
summary = table(c, 600, 107, [
    ["顧客", "合計"],
    ["ひまわり商事", "8,160"],
    ["㈱ひまわり商事", "360"],
    ["ひまわり商事␣", "1,100"],
    ["ヒマワリ商事", "900"],
], [170, 90], color="orange", h=34, fills={(r, 0): "red" for r in range(2, 5)})
c.link(sales, summary, label="SUMIF")
c.text(480, 310, "人には同じ会社に見えても、関数やマクロには別の文字列。1 社が 4 社に分かれる（␣ は末尾の空白）", scale="caption")
c.save("02-name-variants.svg")

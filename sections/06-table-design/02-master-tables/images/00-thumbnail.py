"""6-2: 売上には顧客名だけを書き、住所や担当者は顧客シートに 1 回だけ書いて引く（LINK）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(1040, 360)
c.text(40, 70, "売上（明細）", scale="heading", align="left")
sales = table(c, 40, 100, [
    ["日付", "顧客", "商品", "数量"],
    ["10/02", "あおば設計事務所", "付箋", "5"],
    ["10/06", "みなと物産", "付箋", "10"],
    ["10/07", "ひまわり商事", "ホッチキス", "2"],
    ["10/07", "ひまわり商事", "コピー用紙", "3"],
    ["10/08", "みなと物産", "ノート", "2"],
], [70, 140, 100, 56], color="green", h=32,
    fills={(r, 1): "yellow" for r in range(1, 6)})
c.text(630, 70, "顧客（マスタ）", scale="heading", align="left")
customers = table(c, 630, 132, [
    ["顧客名", "住所", "ご担当者"],
    ["ひまわり商事", "中央区ひまわり 2-4-6", "山田 様"],
    ["あおば設計事務所", "新宿区あおば台 3-1-8", "佐藤 様"],
    ["みなと物産", "横浜市中区みなと 5-5-1", "鈴木 様"],
], [130, 170, 84], color="blue", h=32, fills={(r, 0): "yellow" for r in range(1, 4)})
c.link(sales, customers, label="顧客名で引く", both=False)
c.text(520, 330, "住所や担当者は顧客シートに 1 回だけ書く。売上には「どの顧客か」だけを書く", scale="caption")
c.save("00-thumbnail.svg")

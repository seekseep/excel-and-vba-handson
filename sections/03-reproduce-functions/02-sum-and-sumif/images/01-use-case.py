"""3-2: SUM と SUMIF の使いどころ。10 月の明細から「月の合計」と「顧客ごとの合計」を出す（SPLITTING + SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1000, 440)
c.text(500, 44, "月末の締め：10 月の明細から 2 つの数字を出す", scale="heading")

rows = c.node(130, 240, "10 月の売上明細\n30 行・5 社", emoji_cp="1f4d2")      # 📒

q_total = c.sticky(300, 100, 300, 110, color="blue", weight="thin")
c.text(450, 140, "「10 月は全部でいくら売れた？」", scale="label")
c.text(450, 180, "全部の行を足す → SUM", scale="body")

q_cust = c.sticky(300, 270, 300, 110, color="yellow", weight="thin")
c.text(450, 310, "「ひまわり商事さんに\nいくら請求する？」", scale="label")
c.text(450, 362, "その会社の行だけ足す → SUMIF", scale="body")

report = c.node(840, 155, "売上合計 68,580 円\n店長に月の報告", emoji_cp="1f4c8")   # 📈
invoice = c.node(840, 325, "13,420 円\n請求書に書く（5 章）", emoji_cp="1f9fe")     # 🧾

c.link(rows, q_total)
c.link(rows, q_cust)
c.link(q_total, report)
c.link(q_cust, invoice)
c.save("01-use-case.svg")

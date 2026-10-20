"""5-2 サムネ: 売上からその顧客の行だけを選んで、請求書の明細に書き写す。5 社分くり返す（SOURCE-PATH-GOAL + CYCLE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(900, 300)
sales = c.node(130, 130, "売上シート", emoji_cp="1f4ca")             # 📊
pick = c.node(450, 130, "この顧客の行？", shape="diamond", color="yellow", w=200, h=110, label_scale="label")
invoice = c.node(770, 130, "請求書の明細", emoji_cp="1f4c4")         # 📄
c.link(sales, pick)
c.link(pick, invoice, label="書き写す")
c.link(invoice, sales, label="次の顧客へ（× 5）", primary=False, route="elbow", start="s", end="s")
c.save("00-thumbnail.svg")

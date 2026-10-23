"""6-1（6-3 の 00-thumbnail と同じ図）: 元のデータ（明細とマスタ）と、見せる表（集計と帳票）を分ける（CONTAINER + SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1100, 360)
inp = c.node(100, 180, "入力", emoji_cp="270d", label_scale="label")          # ✍️
c.text(100, inp.below(26), "プルダウンや\nフォームで", scale="caption")
data_box = c.sticky(220, 40, 420, 280, color="green", weight="thin", dash="dashed")
c.text(430, 76, "元のデータ（1 行に 1 件）", scale="label")
detail = c.node(310, 180, "明細", emoji_cp="1f4d2", label_scale="label")      # 📒
c.text(310, detail.below(26), "売上・出欠など\n毎日増える", scale="caption")
master = c.node(550, 180, "マスタ", emoji_cp="1f5c2", label_scale="label")    # 🗂️
c.text(550, master.below(26), "顧客・社員など\nたまに変わる", scale="caption")
c.link(detail, master, label="コードで引く", primary=False, both=True)

out_box = c.sticky(730, 40, 340, 280, color="blue", weight="thin", dash="dashed")
c.text(900, 76, "見せる表（作るもの）", scale="label")
summary = c.node(815, 180, "集計", emoji_cp="1f4ca", label_scale="label")     # 📊
c.text(815, summary.below(26), "関数・ピボット", scale="caption")
report = c.node(985, 180, "帳票", emoji_cp="1f4c4", label_scale="label")      # 📄
c.text(985, report.below(26), "請求書・一覧\nマクロで作る", scale="caption")

c.link(inp, data_box)
c.link(data_box, out_box, label="作る")
c.save("03-input-and-view.svg")

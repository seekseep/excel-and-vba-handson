"""6-3: よくある出欠表の 4 つの問題（①〜④）と、それを1 行に 1 件の出欠と社員マスタに設計し直す（SOURCE-PATH-GOAL + LINK）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(1060, 640)
c.text(40, 36, "よくある出欠表", scale="heading", align="left")
# ④ ほかの研修の表が後ろに重なっている
for d in (16, 8):
    c.raw(f'<rect x="{40 + d}" y="{76 + d}" width="420" height="144" fill="#ffffff" '
          f'stroke="#fdba74" stroke-width="1.5"/>')
before = table(c, 40, 76, [
    ["氏名", "部署", "4/1", "4/2", "4/3", "出席"],
    ["佐藤 花子", "営業部", "◯", "◯", "×", "2"],
    ["鈴木 一郎", "総務部", "◯", "遅刻", "◯", "3"],
    ["田中 次郎", "営業部", "◯", "◯", "◯", "3"],
], [100, 80, 60, 60, 60, 60], color="orange", h=36,
    fills={(2, 3): "red", (1, 5): "yellow", (2, 5): "yellow", (3, 5): "yellow"})

def mark(x, y, n):
    c.raw(f'<circle cx="{x}" cy="{y}" r="15" fill="#ef4444"/>')
    c.text(x, y + 7, n, scale="caption", fill="#ffffff", bold=True)

mark(310, 62, "1")
mark(340, 148, "2")
mark(180, 62, "3")
mark(476, 236, "4")
c.text(530, 110, "① 日付の列が増える\n② 「遅刻」の数え方が決まっていない\n③ 部署を何か所も持っている\n④ 研修ごとに表が分かれている",
       scale="body", align="left")

c.text(40, 370, "1 行に 1 件の出欠", scale="heading", align="left")
att = table(c, 40, 396, [
    ["日付", "社員番号", "研修", "出欠"],
    ["2026/04/01", "E012", "新人研修", "出席"],
    ["2026/04/01", "E027", "新人研修", "出席"],
    ["2026/04/02", "E027", "新人研修", "遅刻"],
    ["2026/04/03", "E012", "新人研修", "欠席"],
    ["…", "…", "…", "…"],
], [110, 90, 100, 70], color="green", h=36)
c.text(700, 370, "社員（マスタ）", scale="heading", align="left")
emp = table(c, 700, 432, [
    ["社員番号", "氏名", "部署"],
    ["E012", "佐藤 花子", "営業部"],
    ["E027", "鈴木 一郎", "総務部"],
    ["E031", "田中 次郎", "営業部"],
], [90, 110, 90], color="blue", h=36)
c.link(att, emp, label="社員番号で引く")
c.link((330, 256), (330, 392), label="設計し直す")
c.save("01-attendance.svg")

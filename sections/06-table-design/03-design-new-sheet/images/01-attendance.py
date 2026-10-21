"""6-3: 研修の出欠表を、1 行に 1 件の出欠と社員マスタに設計し直す（SOURCE-PATH-GOAL + LINK）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(1060, 600)
c.text(40, 50, "よくある出欠表", scale="heading", align="left")
before = table(c, 40, 76, [
    ["氏名", "部署", "4/1", "4/2", "4/3", "出席"],
    ["佐藤 花子", "営業部", "◯", "◯", "×", "2"],
    ["鈴木 一郎", "総務部", "◯", "遅刻", "◯", "3"],
    ["田中 次郎", "営業部", "◯", "◯", "◯", "3"],
], [100, 80, 60, 60, 60, 60], color="orange", h=36,
    fills={(2, 3): "red", (1, 5): "yellow", (2, 5): "yellow", (3, 5): "yellow"})
c.text(500, 110, "・日付が横に増えていく\n・「遅刻」を出席に数える？ 数えない？\n・異動で部署が変わったら全部の表を直す\n・同じ人が別の研修の表にも何度も出てくる",
       scale="body", align="left")

c.text(40, 330, "1 行に 1 件の出欠", scale="heading", align="left")
att = table(c, 40, 356, [
    ["日付", "社員番号", "研修", "出欠"],
    ["2026/04/01", "E012", "新人研修", "出席"],
    ["2026/04/01", "E027", "新人研修", "出席"],
    ["2026/04/02", "E027", "新人研修", "遅刻"],
    ["2026/04/03", "E012", "新人研修", "欠席"],
    ["…", "…", "…", "…"],
], [110, 90, 100, 70], color="green", h=36)
c.text(700, 330, "社員（マスタ）", scale="heading", align="left")
emp = table(c, 700, 392, [
    ["社員番号", "氏名", "部署"],
    ["E012", "佐藤 花子", "営業部"],
    ["E027", "鈴木 一郎", "総務部"],
    ["E031", "田中 次郎", "営業部"],
], [90, 110, 90], color="blue", h=36)
c.link(att, emp, label="社員番号で引く")
c.link((330, 222), (330, 352), label="設計し直す")
c.save("01-attendance.svg")

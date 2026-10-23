"""6-2: 名前でつなぐと、社名の変更で全部の行を直し、同姓同名を区別できない。コードなら 1 セルで済み、区別もできる（BALANCE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(1000, 540)
h = 34
RED, GREEN = "#b91c1c", "#15803d"
c.text(250, 44, "✕ 名前でつなぐ", scale="heading", fill=RED)
c.text(750, 44, "◯ コードでつなぐ", scale="heading", fill=GREEN)
c.raw('<line x1="500" y1="24" x2="500" y2="520" stroke="#cbd5e1" stroke-width="2" stroke-dasharray="7 6"/>')

# 上段: 社名が変わったとき
c.text(250, 92, "社名が「ひまわりホールディングス」に変わった", scale="caption")
table(c, 60, 110, [
    ["日付", "顧客"],
    ["10/07", "ひまわり商事"],
    ["10/13", "ひまわり商事"],
    ["10/15", "ひまわり商事"],
], [80, 300], color="orange", h=h, scale="body", fills={(r, 1): "red" for r in (1, 2, 3)})
c.text(250, 280, "過去の明細も全部の行を直す", scale="caption", fill=RED)

c.text(750, 92, "社名が「ひまわりホールディングス」に変わった", scale="caption")
table(c, 520, 110, [
    ["日付", "顧客コード"],
    ["10/07", "C001"],
    ["10/13", "C001"],
    ["10/15", "C001"],
], [60, 100], color="orange", h=h, scale="body")
table(c, 706, 144, [
    ["顧客コード", "顧客名"],
    ["C001", "ひまわりホールディングス"],
], [90, 200], color="blue", h=h, scale="body", fills={(1, 1): "green"})
c.link((680, 110 + h * 2.5), (706, 144 + h * 1.5), primary=False)
c.text(750, 280, "マスタの 1 セルを直すだけ", scale="caption", fill=GREEN)

# 下段: 同じ名前が 2 人
c.text(250, 344, "社員に「佐藤 花子」さんが 2 人いる", scale="caption")
table(c, 60, 362, [
    ["日付", "氏名"],
    ["4/1", "佐藤 花子"],
], [80, 120], color="orange", h=h, scale="body", fills={(1, 1): "red"})
table(c, 300, 362, [
    ["氏名", "部署"],
    ["佐藤 花子", "営業部"],
    ["佐藤 花子", "総務部"],
], [100, 80], color="blue", h=h, scale="body")
c.link((260, 362 + h * 1.5), (300, 362 + h * 1.5), primary=False, dash="dashed")
c.text(250, 500, "どちらの佐藤さんか区別できない", scale="caption", fill=RED)

c.text(750, 344, "社員に「佐藤 花子」さんが 2 人いる", scale="caption")
table(c, 530, 362, [
    ["日付", "社員番号"],
    ["4/1", "E027"],
], [60, 90], color="orange", h=h, scale="body")
table(c, 730, 362, [
    ["社員番号", "氏名", "部署"],
    ["E012", "佐藤 花子", "営業部"],
    ["E027", "佐藤 花子", "総務部"],
], [80, 90, 70], color="blue", h=h, scale="body", fills={(2, 0): "green"})
c.link((680, 362 + h * 1.5), (730, 362 + h * 2.5), primary=False)
c.text(750, 500, "番号で 1 人に決まる", scale="caption", fill=GREEN)
c.save("05-name-vs-code.svg")

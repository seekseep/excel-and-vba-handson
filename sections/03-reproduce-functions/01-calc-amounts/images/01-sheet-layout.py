"""第3章: 売上シートの構成。A〜F 列が明細、H 列が項目、I 列に関数・J 列に VBA の答え（CONTAINER / PART-WHOLE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(960, 420)
c.text(480, 44, "「売上」シート", scale="heading")

c.sticky(40, 80, 520, 300, color="gray", weight="thin")
c.text(300, 116, "A〜F 列  売上明細", scale="label")
cols = ["日付", "顧客", "商品", "単価", "数量", "金額"]
for i, name in enumerate(cols):
    x = 60 + i * 82
    color = "orange" if name == "金額" else "gray"
    c.sticky(x, 140, 74, 40, color=color, rx=6, weight="thin")
    c.text(x + 37, 166, name, scale="caption")
c.text(300, 230, "2〜31 行目に 30 件", scale="body")
c.text(300, 290, "F 列（金額）は空。関数か VBA で埋める", scale="caption")

c.sticky(600, 80, 320, 300, color="gray", weight="thin")
c.text(760, 116, "H〜J 列  集計欄", scale="label")
c.sticky(620, 140, 90, 40, color="gray", rx=6, weight="thin")
c.text(665, 166, "H 項目", scale="caption")
c.sticky(716, 140, 90, 40, color="yellow", rx=6, weight="thin")
c.text(761, 166, "I 関数", scale="caption")
c.sticky(812, 140, 90, 40, color="blue", rx=6, weight="thin")
c.text(857, 166, "J VBA", scale="caption")
c.text(760, 230, "2 行目  売上合計", scale="body")
c.text(760, 262, "5〜9 行目  顧客別の合計", scale="body")
c.text(760, 320, "I と J が同じなら正解", scale="caption")
c.save("01-sheet-layout.svg")

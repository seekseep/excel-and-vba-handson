"""第4章サムネ: VBA は関数ではできない「操作」をする。シートを作る・飾る・保存する（CENTER-PERIPHERY）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(900, 340)
vba = c.node(450, 170, "VBA", emoji_cp="2699")                         # ⚙️
add = c.node(150, 90, "シートを作る", emoji_cp="1f4c4")                # 📄
paint = c.node(150, 260, "表を飾る", emoji_cp="1f3a8")                 # 🎨
save = c.node(750, 90, "保存する", emoji_cp="1f4be")                   # 💾
split = c.node(750, 260, "振り分ける", emoji_cp="1f5c2")               # 🗂️
for n in (add, paint, save, split):
    c.link(vba, n)
c.save("00-thumbnail.svg")

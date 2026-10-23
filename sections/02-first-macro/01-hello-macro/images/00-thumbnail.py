"""2-1 サムネ: VBE に Sub を書く → F5 で実行 → メッセージの小窓が出る（SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(880, 300)
c.text(440, 52, "書いて、動かして、メッセージを出す", scale="heading")
code = c.node(150, 170, "Sub Hello() を書く", emoji_cp="1f4dd")      # 📝
run = c.node(440, 170, "F5 で実行", emoji_cp="25b6")                 # ▶️
msg = c.node(730, 170, "小窓が出る", emoji_cp="1f4ac")               # 💬
c.link(code, run)
c.link(run, msg)
c.save("00-thumbnail.svg")

"""第3章: デバッガーの使い方。止める → 変数を見る → 1 行進める をくり返し、おかしな値を見つける（CYCLE + SOURCE-PATH-GOAL）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1040, 330)
stop = c.node(130, 140, "ブレークポイント\nで止める", emoji_cp="1f534")   # 🔴
look = c.node(400, 140, "変数の値を見る", emoji_cp="1f50d")                # 🔍
step = c.node(660, 140, "1 行だけ進める", emoji_cp="23ed")                 # ⏭️
bug = c.node(910, 140, "おかしな値を\n見つける", emoji_cp="1f41e")         # 🐞
c.link(stop, look, label="F5")
c.link(look, step, label="F8")
c.link(step, look, label="くり返す", primary=False, route="elbow", start="s", end="s")
c.link(step, bug, dash="dashed")
c.save("00-thumbnail.svg")

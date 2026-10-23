"""サンプルのスプレッドシートは閲覧のみ。「コピーを作成」で自分のマイドライブに作ってから使う。始める状態と完成版の 2 種類（SOURCE-PATH-GOAL + BLOCKAGE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1040, 450)
c.text(190, 40, "サンプル（閲覧のみ）", scale="label")
c.text(850, 40, "自分のコピー（マイドライブ）", scale="label")

rows = [
    (140, "節の冒頭のボタン", "サンプルの\nスプレッドシートをコピー", "始める状態", "コードは入っていない"),
    (330, "完成コードの欄のボタン", "完成版の\nスプレッドシートをコピー", "完成版", "Apps Script 入り"),
]
for cy, where, button, name, note in rows:
    for x, cp, color, status, fill in [(40, "1f512", "gray", "書き込めない", "#b91c1c"),
                                       (700, "270f", "green", "ここで作業する", "#15803d")]:
        box = c.sticky(x, cy - 70, 300, 140, color=color, weight="thin")
        c.emoji(cp, x + 18, cy - 54, 40)                                   # 🔒 / ✏️
        c.text(x + 170, cy - 22, name, scale="heading")
        c.text(x + 150, cy + 18, note, scale="caption")
        c.text(x + 150, cy + 48, status, scale="caption", fill=fill)
        if x == 40:
            src = box
        else:
            dst = box
    c.link(src, dst, label="コピーを作成", label_pos="above")
    c.text(520, cy + 28, f"{where}\n「{button}」", scale="caption")
c.save("01-copy-sample.svg")

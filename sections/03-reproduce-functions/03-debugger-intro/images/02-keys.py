"""3-3: デバッガーの操作の対応表。Excel（VBE）のキー・メニューと、Apps Script エディタのボタン（表そのもの）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
sys.path.insert(0, "../../../../tools")
from genfig import Canvas
from figtable import table

c = Canvas(900, 340)
c.text(450, 44, "同じ操作の、Excel とスプレッドシートでの押し方", scale="heading")
table(c, 40, 76, [
    ["手順", "Excel（VBE）", "スプレッドシート（Apps Script）"],
    ["① 置く", "左の灰色の帯をクリック（F9）", "行番号の左をクリック"],
    ["② 実行する", "マクロの中で F5", "「デバッグ」（「実行」では止まらない）"],
    ["③ 変数を見る", "表示 → ローカル ウィンドウ", "右の「デバッガ」パネルの「変数」"],
    ["④ 1 行進める", "F8（Mac は ⌘ + Shift + I）", "ステップ オーバー"],
    ["続きを一気に", "F5", "再開（▶）"],
    ["⑤ 止める", "■ リセット", "■ 停止"],
], [150, 290, 380], color="blue", h=34, scale="body",
    fills={(0, 1): "green", (0, 2): "yellow"})
c.save("02-keys.svg")

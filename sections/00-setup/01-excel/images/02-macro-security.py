"""ダウンロードしたマクロ入りのブックを動かすまで。Windows は「許可する」→「コンテンツの有効化」、Mac は開くときに有効化（SOURCE-PATH-GOAL + BLOCKAGE）。"""
import sys
sys.path.insert(0, "/Users/seekseep/.claude/skills/genfig")
from genfig import Canvas

c = Canvas(1040, 520)

dl = c.node(92, 250, "ダウンロード", emoji_cp="1f4e5", w=80, h=64, label_scale="label")   # 📥
c.text(92, 316, ".xlsm", scale="caption")

c.text(200, 50, "Windows", scale="heading", align="left")
prop = c.sticky(200, 70, 310, 140, color="red", weight="thin")
c.text(355, 108, "① ファイルのプロパティ", scale="body", fill="#0f172a")
c.text(355, 142, "右クリック →「プロパティ」", scale="caption")
c.text(355, 182, "「許可する」にチェック", scale="label", fill="#b91c1c")
c.text(355, 244, "しないと赤いバー「マクロの実行がブロック…」", scale="caption")

bar = c.sticky(560, 70, 290, 140, color="yellow", weight="thin")
c.text(705, 108, "② 開くと黄色いバー", scale="body", fill="#0f172a")
c.text(705, 142, "セキュリティの警告", scale="caption")
c.text(705, 182, "「コンテンツの有効化」", scale="label", fill="#b45309")

c.text(200, 330, "Mac", scale="heading", align="left")
mac = c.sticky(200, 350, 650, 100, color="yellow", weight="thin")
c.text(525, 388, "開くときに「このブックにはマクロが含まれています」", scale="caption")
c.text(525, 426, "「マクロを有効にする」", scale="label", fill="#b45309")

ok = c.node(960, 250, "マクロが動く", emoji_cp="2705", w=80, h=64, label_scale="label")   # ✅

c.link((dl, "n"), (prop, "w"), route="elbow")
c.link((dl, "s"), (mac, "w"), route="elbow")
c.link(prop, bar)
c.link((bar, "e"), (ok, "n"), route="elbow")
c.link((mac, "e"), (ok, "s"), route="elbow")
c.text(525, 494, "自分で書いた（中身を確かめた）ブックにだけ行う", scale="body", fill="#0f172a")
c.save("02-macro-security.svg")

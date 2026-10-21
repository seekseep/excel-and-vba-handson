"""genfig の図に小さな表を描く補助。6 章の図（表の形そのものが主役の図）で使う。

各節の images/*.py から次のように読み込む。

    sys.path.insert(0, "../../../../tools")
    from figtable import table
"""
from genfig.geom import Poly, pts_rect
from genfig.node import Node
from genfig.text import esc
from genfig.tokens import CENTER_R, PALETTE, SCALE

GRID = "#cbd5e1"


def table(c, x, y, rows, widths, color="blue", h=32, fills=None, scale="caption", vmerges=()):
    """rows の 1 行目を見出しにした表を描き、外枠の Node を返す（link() で結べる）。

    - セルは文字列か (文字列, 横に結合する列数)。結合したセルは左の列の番号で数える
    - fills は {(行, 列): 色名}。見出し以外の行も塗れる。None を渡すと見出しを白にする
    - 文字列が "" のセルは空のまま描く
    - vmerges は [(行, 列, 行数, 文字列)]。縦に結合したセルを重ねて描く
    """
    size, weight = SCALE[scale]
    fills = fills or {}
    border = PALETTE[color]["border"]
    for r, row in enumerate(rows):
        cx, k = x, 0
        for cell in row:
            text, span = (cell, 1) if isinstance(cell, str) else cell
            w = sum(widths[k:k + span])
            name = fills.get((r, k), color if r == 0 else None)
            bg = PALETTE[name]["bg"] if name else "#ffffff"
            c.raw(f'<rect x="{cx:g}" y="{y + r * h:g}" width="{w:g}" height="{h:g}" '
                  f'fill="{bg}" stroke="{GRID}" stroke-width="1"/>')
            if text:
                ty = y + r * h + h / 2 + size * CENTER_R
                fw = "700" if r == 0 else weight
                c.raw(f'<text x="{cx + w / 2:g}" y="{ty:g}" text-anchor="middle" '
                      f'font-size="{size}" font-weight="{fw}" fill="#0f172a">{esc(text)}</text>')
            cx += w
            k += span
    for r, k, n, text in vmerges:
        _cell(c, x + sum(widths[:k]), y + r * h, widths[k], h * n, text, size, weight)
    W, H = sum(widths), h * len(rows)
    c.raw(f'<rect x="{x:g}" y="{y:g}" width="{W:g}" height="{H:g}" fill="none" '
          f'stroke="{border}" stroke-width="2"/>')
    return Node(x + W / 2, y + H / 2, W, H, Poly(pts_rect(x, y, W, H)), color=color)


def _cell(c, x, y, w, h, text, size, weight, bg="#ffffff"):
    c.raw(f'<rect x="{x:g}" y="{y:g}" width="{w:g}" height="{h:g}" '
          f'fill="{bg}" stroke="{GRID}" stroke-width="1"/>')
    c.raw(f'<text x="{x + w / 2:g}" y="{y + h / 2 + size * CENTER_R:g}" text-anchor="middle" '
          f'font-size="{size}" font-weight="{weight}" fill="#0f172a">{esc(text)}</text>')

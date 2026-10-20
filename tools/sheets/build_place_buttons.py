#!/usr/bin/env python3
"""
配布用スプレッドシートにボタンを置く Apps Script（tools/sheets/place_buttons.gs）を作る。

スプレッドシートの図形描画は API から作れないので、ボタンの画像を置いてスクリプト（関数名）を割り当て、
ボタンの代わりにする。画像はこのスクリプトで描き、place_buttons.gs に埋め込む（貼って実行するだけで済むように）。

使い方（保守者のみ。macOS のヒラギノを使う）:
    pip install pillow
    python3 tools/sheets/build_place_buttons.py
できた place_buttons.gs を script.google.com の新しいプロジェクトに貼り、placeButtons を実行する。
"""

import base64
import io
import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).resolve().parent
OUT = HERE / "place_buttons.gs"
FONT = "/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc"

# 配布用スプレッドシートを置いた Google ドライブのフォルダ
FOLDER_ID = "1eYWiXBF5k72-dVUu5QHtjcENpTosdUSI"

# 画面では 120 × 40 px で置く。くっきり見えるよう 2 倍の大きさで描く
SCALE = 2
WIDTH, HEIGHT = 120 * SCALE, 40 * SCALE

# GAS の関数名と表示する文字。tools/build_workbooks.py の Excel のボタンと文字・位置を揃える
LABELS = {
    "calcAmounts": "金額を計算",
    "sumTotal": "売上合計を出す",
    "sumByCustomer": "顧客別に集計",
    "sumAllCustomers": "全顧客を集計",
    "sumTotalWithBug": "合計（バグあり）",
    "showObjects": "ブックの情報を出す",
    "decorateSales": "表を整える",
    "splitByCustomer": "顧客別に分ける",
    "deleteDemoSheets": "増えたシートを消す",
    "saveBackup": "バックアップを保存",
    "createInvoices": "請求書を作る",
    "deleteInvoices": "請求書シートを消す",
}

SALES = [("売上", "L2", "calcAmounts"), ("売上", "L5", "sumTotal"), ("売上", "L8", "sumByCustomer")]
DEBUG = [("売上", "L11", "sumAllCustomers"), ("売上", "L14", "sumTotalWithBug")]
DEMO = [
    ("売上", "L2", "showObjects"), ("売上", "L5", "decorateSales"), ("売上", "L8", "splitByCustomer"),
    ("売上", "L11", "deleteDemoSheets"), ("売上", "L14", "saveBackup"),
]
INVOICE = [("顧客", "F2", "createInvoices")]

# ファイル名の先頭（節の番号）→ 置くボタン。5-3 の「請求書シートを消す」は受講者が自分で置く
PLACEMENTS = {
    "3-1": SALES,
    "3-2": SALES,
    "3-3": SALES,
    "3-4": SALES + DEBUG,
    "4-1": DEMO[:1],
    "4-2": DEMO,
    "5-1": INVOICE,
    "5-2": INVOICE,
    "5-3": INVOICE,
}


def draw(text):
    image = Image.new("RGBA", (WIDTH, HEIGHT), (0, 0, 0, 0))
    d = ImageDraw.Draw(image)
    d.rounded_rectangle(
        (1, 1, WIDTH - 2, HEIGHT - 2), radius=6 * SCALE,
        fill=(255, 255, 255), outline=(26, 115, 232), width=SCALE,
    )
    size = 13 * SCALE
    font = ImageFont.truetype(FONT, size)
    # 長い文字はボタンの幅に収まるまで小さくする
    while d.textlength(text, font=font) > WIDTH - 12 * SCALE:
        size -= 1
        font = ImageFont.truetype(FONT, size)
    d.text((WIDTH / 2, HEIGHT / 2), text, font=font, fill=(26, 115, 232), anchor="mm")
    return image


def png_base64(text):
    buf = io.BytesIO()
    # 色数を減らして小さくする（スクリプトに埋め込むため）
    draw(text).quantize(colors=16).save(buf, format="PNG", optimize=True)
    return base64.b64encode(buf.getvalue()).decode()


GAS = """\
// 配布用スプレッドシートにボタン（スクリプトを割り当てた画像）を置く保守用スクリプト。
// tools/sheets/build_place_buttons.py が作る。手で編集しない。
// フォルダの持ち主のアカウント（info@yokoyamakikaku.work）で script.google.com に新しいプロジェクトを作って貼り、
// placeButtons を実行する（何度実行してもよい）。

const FOLDER_ID = '__FOLDER_ID__';
const WIDTH = __WIDTH__;
const HEIGHT = __HEIGHT__;
const LABELS = __LABELS__;
const PLACEMENTS = __PLACEMENTS__;
const IMAGES = __IMAGES__;

function placeButtons() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const me = Session.getEffectiveUser().getEmail();
  const found = [];
  const iterator = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (iterator.hasNext()) found.push(iterator.next());
  for (let file of found) {
    // 持ち主を実行した人にそろえる（別のアカウントでアップロードしたものは、コピーを作って置き換える）。
    // 元のファイルはフォルダから外すだけなので、外したあとで元の持ち主がゴミ箱に移す。
    if (file.getOwner().getEmail() !== me) {
      const copy = file.makeCopy(file.getName(), folder);
      folder.removeFile(file);
      console.log('持ち主を移した: ' + file.getName() + '（元のファイル ' + file.getId() + ' をフォルダから外した）');
      file = copy;
    }
    const buttons = PLACEMENTS[file.getName().split(' ')[0]];
    if (!buttons) {
      console.log('対象外: ' + file.getName());
      continue;
    }
    const ss = SpreadsheetApp.openById(file.getId());
    // 前に置いたボタン（スクリプトを割り当てた画像）を消してから置き直す
    for (const sheet of ss.getSheets()) {
      for (const image of sheet.getImages()) {
        if (image.getScript()) image.remove();
      }
    }
    for (const [sheetName, cell, fn] of buttons) {
      const sheet = ss.getSheetByName(sheetName);
      const range = sheet.getRange(cell);
      const blob = Utilities.newBlob(Utilities.base64Decode(IMAGES[fn]), 'image/png', fn + '.png');
      const image = sheet.insertImage(blob, range.getColumn(), range.getRow(), 4, 3);
      image.setWidth(WIDTH).setHeight(HEIGHT);
      image.setAltTextTitle(LABELS[fn]);
      image.assignScript(fn);
    }
    console.log('置いた: ' + file.getName() + '（' + buttons.length + ' 個）');
  }
}
"""


def main():
    images = {name: png_base64(text) for name, text in LABELS.items()}
    placements = {k: [list(b) for b in v] for k, v in PLACEMENTS.items()}
    text = (
        GAS.replace("__FOLDER_ID__", FOLDER_ID)
        .replace("__WIDTH__", str(WIDTH // SCALE))
        .replace("__HEIGHT__", str(HEIGHT // SCALE))
        .replace("__LABELS__", json.dumps(LABELS, ensure_ascii=False, indent=2))
        .replace("__PLACEMENTS__", json.dumps(placements, ensure_ascii=False))
        .replace("__IMAGES__", json.dumps(images, indent=2))
    )
    OUT.write_text(text)
    print(f"[build_place_buttons] {OUT.relative_to(HERE.parent.parent)}")


if __name__ == "__main__":
    main()

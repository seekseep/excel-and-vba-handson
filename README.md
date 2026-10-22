# はじめての VBA ハンズオン

VBA 初心者向けの勉強会「Excel 作業を自動化しよう」の教材です。Excel（VBA）と Google スプレッドシート（Apps Script）の両方に対応しています。

- 教材サイト: <https://seekseep.github.io/excel-and-vba-handson/>
- 教材の本文: [sections/](./sections/)（講義の最初に使う紹介は [sections/INTRODUCTION.md](./sections/INTRODUCTION.md)）

## 構成

| ディレクトリ | 中身 |
|---|---|
| `sections/` | 教材の本文・図・完成コード・練習用ブック |
| `site/` | 解説サイト（`sections/` から生成する Astro Starlight） |
| `tools/` | 練習用ブックの生成など、保守用のスクリプト |

## サイトをローカルで動かす

```sh
cd site
pnpm install
pnpm run dev
```

練習用ブックを作り直すときは `python3 tools/build_workbooks.py`（`openpyxl` が必要）。

## ライセンス

[CC0 1.0](https://creativecommons.org/publicdomain/zero/1.0/deed.ja)。ただし図の中の絵文字は [Twemoji](https://github.com/jdecked/twemoji)（CC BY 4.0）です。登場する会社・人物・住所はすべて架空のものです。

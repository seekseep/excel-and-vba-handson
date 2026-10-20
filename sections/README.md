---
docs: true
title: 進め方
sidebar:
  label: 進め方
  order: 0
---

# 進め方

## 章と節

教材は 5 つの **章** でできていて、講義の章とハンズオンの章が交互に並びます。章はいくつかの **節** に分かれています。

題材は架空の文房具店「そよかぜ文具店」の、月末の請求書づくりです。30 行の売上明細から顧客ごとの請求額を出し、5 社分の請求書を PDF にするまでを、手作業 → Excel 関数 → VBA と道具を持ち替えながら自動化していきます。第 3 章と第 5 章は同じ売上データを使います。

| 章 | 形式 | 中身 | 完成条件 |
|---|---|---|---|
| [1. そのExcel作業、本当に手作業が必要？](./01-why-automate/01-monthly-closing/LECTURE.md) | 講義 | 今日の課題、関数と VBA の違い | — |
| [2. はじめてのマクロを動かそう](./02-first-macro/01-hello-macro/LECTURE.md) | 講義＋実習 | VBE、Sub、セルの読み書き、変数 | 自分で書いたマクロが動く |
| [3. Excel関数をVBAで再現しよう](./03-reproduce-functions/01-calc-amounts/LECTURE.md) | ハンズオン（目標・小） | For / If、SUM と SUMIF の再現、デバッガー | VBA で集計した結果をセルに出せる |
| [4. VBAだからできること](./04-what-vba-can-do/01-object-model/LECTURE.md) | 講義＋デモ | オブジェクト、シート操作、書式、保存 | — |
| [5. ボタンひとつで請求書を作ろう](./05-invoice-button/01-copy-template/LECTURE.md) | ハンズオン（目標・中） | テンプレートへの転記、PDF、ボタン | ボタンを押すと請求書ができる |

節ごとに、練習用のブックとその節までの完成コードを入れた ZIP があります。同じ章の節は同じブックで続けて進めるので、ダウンロードは章の最初の節だけで大丈夫です（途中から始めるときや答え合わせに、ほかの節の ZIP を使ってください）。

## 目標の大きさ

ハンズオンには 3 段階の目標があります。

| 目標 | 中身 | 位置づけ |
|---|---|---|
| **小** | Excel 関数でやっていた計算を VBA で再現する | 全員がここまで完成させる |
| **中** | シートを作り、帳票を出す実用的な自動化 | 勉強会の中で取り組む |
| **大** | 外部データや Web API とつなぐ | 発展・持ち帰り |

大きな課題の途中で止まるより、小さな課題を最後まで動かしたほうが得るものは大きいです。迷ったら「小」を確実に終わらせてください。

## Excel とスプレッドシート

教材は Excel（VBA）が主ですが、Google スプレッドシート（Apps Script、以下 GAS）でも同じ課題に取り組めます。手順やコードが分かれる箇所には、次のような切り替えタブがあります。

:::::platform
::::excel
Excel を使う人向けの手順と VBA のコードが、ここに出ます。
::::
::::sheets
スプレッドシートを使う人向けの手順と GAS のコードが、ここに出ます。
::::
:::::

どちらかのタブを押すと、すべてのページのタブがそろって切り替わります（選んだ環境はブラウザが覚えておきます）。

- 会社や学校の PC で Excel のマクロが禁止されている
- Mac で Excel を持っていない

という場合は、スプレッドシートで参加してください。VBA と GAS の書き方の違いは [VBA と GAS の対応表](./CHEATSHEET.md) にまとめてあります。

## 準備

- [Excel の準備](./EXCEL.md)
- [スプレッドシートの準備](./SHEETS.md)

## 触れる VBA

| 分類 | 項目 | 章 |
|---|---|---|
| 実行環境 | 開発タブ、VBE、標準モジュール、マクロの実行、.xlsm | 2 |
| 基本構文 | `Sub`、`Dim` と型、四則演算 | 2 |
| 基本構文 | `For / Next`、`If / Then / Else` | 3 |
| 基本構文 | `For Each / Next` | 4・5 |
| Excel の操作 | `Range`、`Cells`、`Value` | 2・3 |
| Excel の操作 | `Worksheet(s)`、`Workbook`、シートの作成・コピー・削除、保存 | 4・5 |
| Excel の操作 | 背景色・文字色・罫線 | 4 |
| イベント | ボタンからマクロを実行する | 5 |
| デバッグ | ブレークポイント、ステップ実行、ローカル ウィンドウ、イミディエイト ウィンドウ、`Debug.Print` | 3 |

`Function`、配列、エラー処理は扱いません。

## 一覧

<!-- 節を追加したらここにリンクを足す -->

### 1. そのExcel作業、本当に手作業が必要？

- [1-1 そのExcel作業、本当に手作業が必要？](./01-why-automate/01-monthly-closing/LECTURE.md)

### 2. はじめてのマクロを動かそう

- [2-1 はじめてのマクロ](./02-first-macro/01-hello-macro/LECTURE.md)
- [2-2 セルと変数](./02-first-macro/02-cells-and-variables/LECTURE.md)

### 3. Excel関数をVBAで再現しよう

- [3-1 金額を計算する](./03-reproduce-functions/01-calc-amounts/LECTURE.md)
- [3-2 SUM と SUMIF を再現する](./03-reproduce-functions/02-sum-and-sumif/LECTURE.md)
- [3-3 デバッガーで原因を探す](./03-reproduce-functions/03-debugger/LECTURE.md)

### 4. VBAだからできること

- [4-1 Excel は部品の入れ子](./04-what-vba-can-do/01-object-model/LECTURE.md)
- [4-2 シートとファイルを操作する](./04-what-vba-can-do/02-sheet-operations/LECTURE.md)

### 5. ボタンひとつで請求書を作ろう

- [5-1 テンプレートをコピーして宛名を書く](./05-invoice-button/01-copy-template/LECTURE.md)
- [5-2 明細を書き写す](./05-invoice-button/02-transfer-details/LECTURE.md)
- [5-3 PDF にしてボタンから実行する](./05-invoice-button/03-pdf-and-button/LECTURE.md)

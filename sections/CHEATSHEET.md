---
docs: true
title: VBA と GAS の対応表
sidebar:
  label: VBA と GAS の対応表
  order: 3
---

# VBA と GAS の対応表

この教材に出てくる書き方を、Excel（VBA）と Google スプレッドシート（GAS）で並べました。考え方は同じで、書き方が違うだけです。

## 基本の形

| やりたいこと | VBA | GAS |
|---|---|---|
| マクロを定義する | `Sub Hello()` 〜 `End Sub` | `function hello() { }` |
| 変数を作る | `Dim total As Long` | `let total = 0;`（あとで変える）／ `const sheet = ...;`（変えない） |
| 文字列 | `"こんにちは"` | `'こんにちは'` |
| 文字列をつなぐ | `"請求書_" & name` | `'請求書_' + name` |
| 等しい | `If a = b Then` | `if (a === b) {` |
| 等しくない | `a <> b` | `a !== b` |
| 条件分岐 | `If 〜 Then` / `Else` / `End If` | `if (〜) { } else { }` |
| 繰り返し | `For i = 2 To lastRow` 〜 `Next i` | `for (let i = 2; i <= lastRow; i++) { }` |
| 全部について繰り返す | `For Each ws In Worksheets` 〜 `Next ws` | `for (const sheet of ss.getSheets()) { }` |
| 別のマクロを呼ぶ | `DeleteInvoices` | `deleteInvoices();` |
| コメント | `' コメント` | `// コメント` |

GAS は大文字と小文字を区別します（`sheet` と `Sheet` は別物）。VBA は区別しません。

## ブック・シート・セル

| やりたいこと | VBA | GAS |
|---|---|---|
| このブック | `ThisWorkbook` | `SpreadsheetApp.getActiveSpreadsheet()` |
| 名前でシートを取る | `Worksheets("売上")` | `ss.getSheetByName('売上')` |
| いま開いているシート | `ActiveSheet` | `SpreadsheetApp.getActiveSheet()` |
| シートを変数に入れる | `Set salesSheet = Worksheets("売上")` | `const salesSheet = ss.getSheetByName('売上');` |
| セルを番地で指す | `Range("B2")` | `sheet.getRange('B2')` |
| セルを行・列の番号で指す | `Cells(2, 6)` | `sheet.getRange(2, 6)` |
| 値を読む | `x = Range("B2").Value` | `x = sheet.getRange('B2').getValue()` |
| 値を書く | `Range("B2").Value = 100` | `sheet.getRange('B2').setValue(100)` |
| データのある最後の行 | `Cells(Rows.Count, 1).End(xlUp).Row` | `sheet.getLastRow()` |
| シートを追加する | `Worksheets.Add` | `ss.insertSheet('名前')` |
| シートをコピーする | `ws.Copy After:=...` | `sheet.copyTo(ss)` |
| シートを消す | `ws.Delete` | `ss.deleteSheet(sheet)` |
| シートの名前を変える | `ws.Name = "請求書"` | `sheet.setName('請求書')` |

VBA の `Range("B2")` のようにシートを書かないと「いま開いているシート」が対象になります。GAS は必ずシートを取ってからセルを指します。

## 書式

| やりたいこと | VBA | GAS |
|---|---|---|
| 背景色 | `.Interior.Color = RGB(31, 78, 121)` | `.setBackground('#1f4e79')` |
| 文字色 | `.Font.Color = RGB(192, 0, 0)` | `.setFontColor('#c00000')` |
| 太字 | `.Font.Bold = True` | `.setFontWeight('bold')` |
| 罫線 | `.Borders.LineStyle = xlContinuous` | `.setBorder(true, true, true, true, true, true)` |

## 画面とデバッグ

| やりたいこと | VBA | GAS |
|---|---|---|
| メッセージを出す | `MsgBox "完了"` | `SpreadsheetApp.getUi().alert('完了')` |
| ログに出す | `Debug.Print total`（イミディエイトウィンドウ） | `console.log(total)`（実行ログ） |
| ブレークポイント | 行の左の灰色の帯をクリック（`F9`） | 行番号の左をクリック |
| 1 行ずつ実行 | `F8`（Mac は `⌘` + `Shift` + `I`） | デバッガーの「ステップ オーバー」 |
| ボタンから実行 | フォームコントロールのボタン →「マクロの登録」 | 図形描画 →「スクリプトを割り当て」 |
| 開いたときに実行 | `Workbook_Open`（ThisWorkbook に書く） | `function onOpen() { }` |

## ファイル

| やりたいこと | VBA | GAS |
|---|---|---|
| ブックの保存場所 | `ThisWorkbook.Path` | `DriveApp.getFileById(ss.getId()).getParents().next()` |
| コピーを保存する | `ThisWorkbook.SaveCopyAs fileName` | `DriveApp.getFileById(ss.getId()).makeCopy(name)` |
| シートを PDF にする | `ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=...` | 専用の命令は無い（[5-3](./05-invoice-button/03-pdf-and-button/LECTURE.md) の `savePdf_` を参照） |

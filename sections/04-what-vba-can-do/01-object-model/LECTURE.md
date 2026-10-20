---
docs: true
title: Excel は部品の入れ子
---


# Excel は部品の入れ子

![Excel の中にブック、ブックの中にシート、シートの中にセル](./images/00-thumbnail.svg)

第 3 章の集計は関数で十分でした。第 4 章では、**関数ではできない操作** を VBA でやってみせます。その前に、VBA から Excel がどう見えているかを整理します。

手元で一緒に動かしたい人は、練習用ブックとデモのコードをダウンロードしてください（見ているだけでも大丈夫です）。

:::download
[練習用ブックとデモのコードをダウンロード](./project.zip)
:::

`data/sales.xlsx` は第 3 章と同じ売上明細で、金額（F 列）まで埋まっています。

## Excel は部品の入れ子でできている

VBA から見ると、Excel は「部品の中に部品が入っている」構造になっています。

| 部品 | VBA での呼び方 | 例 |
|---|---|---|
| Excel そのもの | `Application` | |
| ブック（ファイル） | `Workbook` | `ThisWorkbook`（このマクロが入っているブック） |
| シート | `Worksheet` | `Worksheets("売上")` |
| セル・範囲 | `Range` | `Range("B2")`、`Cells(2, 2)` |

この部品のことを **オブジェクト** と呼びます。外側から内側へ、`.`（ドット）でたどって指定します。

```vb
ThisWorkbook.Worksheets("売上").Range("B2").Value
```

「このブックの」「売上シートの」「B2 セルの」「値」と読みます。第 2〜3 章で書いた `Range("B2")` は、ブックとシートを省略した書き方で、そのときは **いま開いているシート** が対象になります。

オブジェクトには 2 種類の要素があります。

- **プロパティ** … 部品の状態。読んだり書き換えたりできる（`.Value` 値、`.Name` 名前、`.Interior.Color` 背景色）
- **メソッド** … 部品にさせる動作（`.Copy` コピー、`.Delete` 削除、`.Add` 追加）

:::::platform
::::excel
### デモ: 部品をたどってみる（`ShowObjects`）

`Debug.Print` で、ブックの名前やシートの数をイミディエイト ウィンドウ（`Ctrl` + `G`）に書き出します。

```vb
Debug.Print "ブックの名前: " & ThisWorkbook.Name
Debug.Print "シートの数: " & ThisWorkbook.Worksheets.Count
Debug.Print "売上!B2 の値: " & ThisWorkbook.Worksheets("売上").Range("B2").Value
```
::::
::::sheets
### デモ: 部品をたどってみる（`showObjects`）

スプレッドシートも同じ入れ子です（スプレッドシート → シート → セル）。`console.log` で実行ログに書き出します。

```js
const ss = SpreadsheetApp.getActiveSpreadsheet();
console.log('スプレッドシートの名前: ' + ss.getName());
console.log('シートの数: ' + ss.getSheets().length);
console.log('売上!B2 の値: ' + ss.getSheetByName('売上').getRange('B2').getValue());
```
::::
:::::

## シートを変数に入れる（Set）

同じシートを何度も使うときは、変数に入れておくと短く書けます。

:::::platform
::::excel
```vb
Dim salesSheet As Worksheet
Set salesSheet = Worksheets("売上")

salesSheet.Range("A1").Value = "日付"
```

数や文字列ではなく **オブジェクトを変数に入れるときは、先頭に `Set` を付けます**。付け忘れると「オブジェクト変数または With ブロック変数が設定されていません」というエラーになります。よく出るエラーなので覚えておいてください。
::::
::::sheets
```js
const salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');

salesSheet.getRange('A1').setValue('日付');
```

GAS では、数や文字列と同じように `const` で入れるだけです。VBA のような `Set` はありません。
::::
:::::

## 次へ

[シートとファイルを操作する](../02-sheet-operations/LECTURE.md) で、表を飾る・シートを作る・保存するデモを見ます。

::codeview[この節のデモのコード]

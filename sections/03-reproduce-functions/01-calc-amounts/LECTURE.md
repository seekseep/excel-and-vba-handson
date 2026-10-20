---
docs: true
title: 金額を計算する
---


# 金額を計算する

![同じ集計を、関数と VBA の 2 通りで出して見比べる](./images/00-thumbnail.svg)

そよかぜ文具店の 10 月の売上明細を集計します。まず Excel 関数で答えを出し、**同じ答えを VBA でも出します**。関数で十分な処理をあえて VBA で書くのは、プログラミングの基本（くり返しと条件分岐）を、答えの分かっている問題で練習するためです。

この節では、関数で答えを用意したあと、各行の金額（単価 × 数量）を VBA で計算します。

## 準備

:::download
[練習用ブックと完成コードをダウンロード](./project.zip)
:::

ZIP を展開すると `data/sales.xlsx` があります。

:::::platform
::::excel
1. `data/sales.xlsx` を Excel で開く（保護ビューの黄色いバーが出たら「編集を有効にする」）
2. すぐに「名前を付けて保存」で **Excel マクロ有効ブック（\*.xlsm）** として保存する
3. `Alt` + `F11` で VBE を開き、「挿入」→「標準モジュール」
::::
::::sheets
1. `data/sales.xlsx` を Google ドライブにアップロードして開く
2. 「ファイル」→ **Google スプレッドシートとして保存**
3. 変換したスプレッドシートで「拡張機能」→「Apps Script」
::::
:::::

「売上」シートの中身はこうなっています。

![売上シートの構成。A〜F 列が明細、H〜J 列が集計欄](./images/01-sheet-layout.svg)

- **A〜E 列** … 30 行の売上明細（2〜31 行目）。F 列「金額」は空
- **H〜J 列** … 集計欄。黄色い **I 列に関数で**、その隣の **J 列に VBA で** 答えを出して見比べます

## まずは関数で解く

VBA の前に、関数で答えを出しておきます。これが VBA の答え合わせになります。

| セル | 入れる式 | 意味 |
|---|---|---|
| F2 | `=D2*E2` | 金額 = 単価 × 数量。F31 までコピーする |
| I2 | `=SUM(F2:F31)` | 売上合計 |
| I5 | `=SUMIF($B$2:$B$31, H5, $F$2:$F$31)` | H5 の顧客の合計。I9 までコピーする |

I2 が **68,580**、I5（ひまわり商事）が **13,420** になれば正解です。

ここまでで集計は終わりました。ここからは、同じことを VBA にやらせます。

## 1 行分の金額を VBA で出す

いきなり全部ではなく、2 行目だけをやってみます。

:::::platform
::::excel
```vb
Sub CalcAmounts()
    Cells(2, 6).Value = Cells(2, 4).Value * Cells(2, 5).Value
End Sub
```

`Cells(2, 6)` は 2 行目の 6 列目、つまり F2 です。`Cells(2, 4)`（D2: 単価）と `Cells(2, 5)`（E2: 数量）を掛けて入れています。
::::
::::sheets
```js
function calcAmounts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  const amount = sheet.getRange(2, 4).getValue() * sheet.getRange(2, 5).getValue();
  sheet.getRange(2, 6).setValue(amount);
}
```

`getSheetByName('売上')` で「売上」シートを名前で取ってきます。`getRange(2, 6)` は 2 行目の 6 列目、つまり F2 です。
::::
:::::

実行すると、F2 の式 `=D2*E2` が消えて、値の `1250` に置き換わります。

:::notice[VBA が書くのは「値」]
関数の式は「どう計算したか」がセルに残りますが、VBA が書いた値には残りません。D2 を書き換えても F2 は変わらず、もう一度マクロを実行するまで古いままです。関数と VBA の大きな違いの 1 つです。
:::

## くり返す（For）

3 行目、4 行目…と 31 行目まで同じことをしたいので、行番号を変数 `i` にして **くり返し** ます。

:::::platform
::::excel
:::code[CalcAmounts を書き換える]{filepath=Module1}
```vb
Sub CalcAmounts()
    Dim i As Long

    For i = 2 To 31
        Cells(i, 6).Value = Cells(i, 4).Value * Cells(i, 5).Value
    Next i
End Sub
```
:::

- `For i = 2 To 31` 〜 `Next i` … `i` を 2, 3, 4, … 31 と 1 ずつ増やしながら、間の行を 30 回くり返す
- `Cells(i, 6)` … `i` 行目の F 列。`i` が 2 のときは F2、3 のときは F3
::::
::::sheets
:::code[calcAmounts を書き換える]{filepath=コード.gs}
```js
function calcAmounts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  for (let i = 2; i <= 31; i++) {
    const amount = sheet.getRange(i, 4).getValue() * sheet.getRange(i, 5).getValue();
    sheet.getRange(i, 6).setValue(amount);
  }
}
```
:::

- `for (let i = 2; i <= 31; i++) { }` … `i` を 2 から始め、31 以下のあいだ、1 ずつ増やしながら `{ }` の中をくり返す
- `getRange(i, 6)` … `i` 行目の F 列
::::
:::::

F2〜F31 がすべて値で埋まり、I2 の SUM が 68,580 のままなら成功です。

### 最後の行を自動で探す

`31` と書いてしまうと、11 月に明細が 40 行になったときに直す必要があります。データのある最後の行を探す書き方に変えましょう。

:::::platform
::::excel
```vb
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastRow
```

`Cells(Rows.Count, 1)` はシートのいちばん下の A 列のセルです。そこから `.End(xlUp)`（`Ctrl` + `↑` と同じ）で上に飛ぶと、データのある最後のセルに止まります。その `.Row`（行番号）が `lastRow` です。
::::
::::sheets
```js
const lastRow = sheet.getLastRow();
for (let i = 2; i <= lastRow; i++) {
```

`getLastRow()` は、データのある最後の行番号を返します。
::::
:::::

## 次へ

[SUM と SUMIF を再現する](../02-sum-and-sumif/LECTURE.md) で、合計と顧客別の合計を VBA で出します。

::codeview[この節の完成コード]{defaultFile="vba/SalesSummary.bas"}

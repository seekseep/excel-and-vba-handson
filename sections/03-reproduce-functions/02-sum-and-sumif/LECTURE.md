---
docs: true
title: SUM と SUMIF を再現する
---


# SUM と SUMIF を再現する

![0 から始めて 1 行ずつ足す。条件に合う行だけ足す](./images/00-thumbnail.svg)

3-1 の続きです。3-1 のブックとモジュールをそのまま使います（ここから始める人は、ZIP の `data/sales.xlsx` で完成コードの `CalcAmounts` を先に実行して F 列を埋めてください）。

**この章の完成条件:** VBA で集計した売上合計と顧客別の合計を、セルに出せる（関数の答えと一致する）。この節で達成します。

## 合計を出す（SUM の再現）

合計は「0 から始めて、1 行ずつ足していく」と考えます。電卓で足し算するときと同じです。

:::::platform
::::excel
```vb
Sub SumTotal()
    Dim lastRow As Long
    Dim i As Long
    Dim total As Long

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    total = 0
    For i = 2 To lastRow
        total = total + Cells(i, 6).Value
    Next i
    Range("J2").Value = total
End Sub
```

`total = total + Cells(i, 6).Value` は「いまの `total` に `i` 行目の金額を足して、`total` に入れ直す」という意味です。
::::
::::sheets
```js
function sumTotal() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  const lastRow = sheet.getLastRow();
  let total = 0;
  for (let i = 2; i <= lastRow; i++) {
    total = total + sheet.getRange(i, 6).getValue();
  }
  sheet.getRange('J2').setValue(total);
}
```

`total = total + ...` は「いまの `total` に `i` 行目の金額を足して、`total` に入れ直す」という意味です。中身を書き換えるので `const` ではなく `let` で作ります。
::::
:::::

J2 が I2 と同じ **68,580** になれば成功です。

## 条件で分ける（SUMIF の再現）

顧客別の合計は「顧客名が一致する行だけ足す」です。ここで **条件分岐** を使います。

:::::platform
::::excel
```vb
Sub SumByCustomer()
    Dim lastRow As Long
    Dim i As Long
    Dim customerName As String
    Dim total As Long

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    customerName = Range("H5").Value
    total = 0
    For i = 2 To lastRow
        If Cells(i, 2).Value = customerName Then
            total = total + Cells(i, 6).Value
        End If
    Next i
    Range("J5").Value = total
End Sub
```

- `If 条件 Then` 〜 `End If` … 条件が成り立つときだけ、間の行を実行する
- `Cells(i, 2).Value = customerName` … `i` 行目の顧客名（B 列）が H5 と同じか。`If` の中の `=` は「等しいか」の意味になります
::::
::::sheets
```js
function sumByCustomer() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  const lastRow = sheet.getLastRow();
  const customerName = sheet.getRange('H5').getValue();
  let total = 0;
  for (let i = 2; i <= lastRow; i++) {
    if (sheet.getRange(i, 2).getValue() === customerName) {
      total = total + sheet.getRange(i, 6).getValue();
    }
  }
  sheet.getRange('J5').setValue(total);
}
```

- `if (条件) { }` … 条件が成り立つときだけ `{ }` の中を実行する
- `===` … 左右が等しいか。`=`（入れる）と区別するために 3 つ並べます
::::
:::::

J5 が I5 と同じ **13,420** になれば成功です。H5 を「みなと物産」に書き換えて実行すると、20,900 になります。

:::notice[条件が成り立たないときもしたいなら]
`Else`（GAS は `else`）で「それ以外のとき」を書けます。たとえば金額が 5,000 以上なら「大口」、それ以外なら「通常」と G 列に書く、などです。

```vb
If Cells(i, 6).Value >= 5000 Then
    Cells(i, 7).Value = "大口"
Else
    Cells(i, 7).Value = "通常"
End If
```
:::

## 次へ

[デバッガーで原因を探す](../03-debugger/LECTURE.md) で、思ったとおりに動かないコードの調べ方を練習します。

::codeview[この節の完成コード]{defaultFile="vba/SalesSummary.bas"}

---
docs: true
title: 明細を書き写す
---


# 明細を書き写す

![売上シートからその顧客の行だけを選び、請求書の明細に書き写す](./images/00-thumbnail.svg)

5-1 の続きです。5-1 のブックとモジュールをそのまま使います（ここから始める人は ZIP の `data/invoice.xlsx` と、5-1 の完成コードから始めてください）。

売上シートからその顧客の明細だけを請求書に書き写し、1 社分を 5 社分に広げます。

## 明細を書き写す

売上シートを上から順に見て、顧客名が一致する行だけを請求書の 12 行目から書き写します。第 3 章の `SumByCustomer` とほぼ同じ形です。違うのは、足すだけでなく **書き写す** ことと、書き写し先の行 `lineRow` を 1 つずつ下げていくことです。

**書く場所:** 5-1 の「宛名と発行日」のあと、`End Sub`（GAS は最後の `}`）の前

:::::platform
::::excel
```vb
    ' 売上からこの顧客の行だけを明細に転記する
    lineRow = 12
    total = 0
    For i = 2 To lastSalesRow
        If salesSheet.Cells(i, 2).Value = customerName Then
            invoiceSheet.Cells(lineRow, 1).Value = salesSheet.Cells(i, 1).Value
            invoiceSheet.Cells(lineRow, 2).Value = salesSheet.Cells(i, 3).Value
            invoiceSheet.Cells(lineRow, 3).Value = salesSheet.Cells(i, 4).Value
            invoiceSheet.Cells(lineRow, 4).Value = salesSheet.Cells(i, 5).Value
            invoiceSheet.Cells(lineRow, 5).Value = salesSheet.Cells(i, 6).Value
            total = total + salesSheet.Cells(i, 6).Value
            lineRow = lineRow + 1
        End If
    Next i
    invoiceSheet.Range("E33").Value = total
    invoiceSheet.Range("B9").Value = total
```

新しい変数を使うので、`Sub` の先頭の `Dim` に次を足し、`lastSalesRow` を求める行も足します。

```vb
    Dim lastSalesRow As Long
    Dim i As Long
    Dim lineRow As Long
    Dim total As Long
```

```vb
    lastSalesRow = salesSheet.Cells(salesSheet.Rows.Count, 1).End(xlUp).Row
```
::::
::::sheets
```js
  // 売上からこの顧客の行だけを明細に転記する
  const lastSalesRow = salesSheet.getLastRow();
  let lineRow = 12;
  let total = 0;
  for (let i = 2; i <= lastSalesRow; i++) {
    if (salesSheet.getRange(i, 2).getValue() === customerName) {
      invoiceSheet.getRange(lineRow, 1).setValue(salesSheet.getRange(i, 1).getValue());
      invoiceSheet.getRange(lineRow, 2).setValue(salesSheet.getRange(i, 3).getValue());
      invoiceSheet.getRange(lineRow, 3).setValue(salesSheet.getRange(i, 4).getValue());
      invoiceSheet.getRange(lineRow, 4).setValue(salesSheet.getRange(i, 5).getValue());
      invoiceSheet.getRange(lineRow, 5).setValue(salesSheet.getRange(i, 6).getValue());
      total = total + salesSheet.getRange(i, 6).getValue();
      lineRow = lineRow + 1;
    }
  }
  invoiceSheet.getRange('E33').setValue(total);
  invoiceSheet.getRange('B9').setValue(total);
```
::::
:::::

売上シートの列（A 日付・C 商品・D 単価・E 数量・F 金額）と、請求書の列（A 日付・B 商品・C 単価・D 数量・E 金額）はずれています。顧客名（売上の B 列）は請求書の明細には書かないからです。

ひまわり商事の請求書に 7 行の明細が入り、ご請求金額が **¥13,420** になれば成功です。第 3 章の顧客別合計と一致しています。

:::notice[うまくいかないときはデバッガー]
明細が 1 行も入らないときは、`If` の行にブレークポイントを置いて、`salesSheet.Cells(i, 2).Value` と `customerName` を見比べてください。顧客名の前後に余計なスペースがあると「一致しない」と判定されます。
:::

## 全顧客分をくり返す

`c = 2` で 1 社だけにしていたところを、顧客シートの 2 行目から最後の行までの **くり返し** に変えます。

:::::platform
::::excel
```vb
    lastCustomerRow = customerSheet.Cells(customerSheet.Rows.Count, 1).End(xlUp).Row

    For c = 2 To lastCustomerRow
        customerName = customerSheet.Cells(c, 1).Value
        ' …5-1 とこの節で書いた、コピー〜合計までの行…
    Next c
```

`Dim lastCustomerRow As Long` も足してください。`For c` の中に、明細をくり返す `For i` が入る形になります（第 3 章のチャレンジと同じ「くり返しの中のくり返し」です）。
::::
::::sheets
```js
  const lastCustomerRow = customerSheet.getLastRow();

  for (let c = 2; c <= lastCustomerRow; c++) {
    const customerName = customerSheet.getRange(c, 1).getValue();
    // …5-1 とこの節で書いた、コピー〜合計までの行…
  }
```

`for (let c ...)` の中に、明細をくり返す `for (let i ...)` が入る形になります（第 3 章のチャレンジと同じ「くり返しの中のくり返し」です）。
::::
:::::

実行すると、5 社分の請求書シートができます。

### 何度実行しても大丈夫にする

もう一度実行すると「その名前はすでに使用されています」というエラーになります。そこで、作る前に前回の請求書シートを消す `DeleteInvoices` を用意して、最初に呼びます。

:::::platform
::::excel
```vb
Sub DeleteInvoices()
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If Left(ws.Name, 4) = "請求書_" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub
```

`Left(ws.Name, 4)` はシート名の左から 4 文字です。「請求書テンプレート」は消さずに、「請求書_」で始まるシートだけを消します。

`CreateInvoices` の `Dim` のあとに `DeleteInvoices` と 1 行書くと、そこでこのマクロが呼ばれます。
::::
::::sheets
```js
function deleteInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const sheet of ss.getSheets()) {
    if (sheet.getName().startsWith('請求書_')) {
      ss.deleteSheet(sheet);
    }
  }
}
```

`startsWith('請求書_')` は「請求書_ で始まるか」です。「請求書テンプレート」は消さずに、「請求書_」で始まるシートだけを消します。

`createInvoices` の最初に `deleteInvoices();` と 1 行書くと、そこでこの関数が呼ばれます。
::::
:::::

## 次へ

[PDF にしてボタンから実行する](../03-pdf-and-button/LECTURE.md) で、仕上げます。

::codeview[この節の完成コード]

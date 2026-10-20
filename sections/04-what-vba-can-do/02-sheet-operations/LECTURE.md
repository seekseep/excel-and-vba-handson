---
docs: true
title: シートとファイルを操作する
---


# シートとファイルを操作する

![VBA は関数ではできない「操作」をする。シートを作る・飾る・保存する](./images/00-thumbnail.svg)

4-1 の続きです。シートを作る、表を飾る、ファイルを保存する。第 5 章の請求書づくりに必要な部品を、講師のデモで一通り見ていきます。

:::download
[練習用ブックとデモのコードをダウンロード](./project.zip)
:::

## デモ 1: 表を飾る（`DecorateSales`）

見出しに色を付け、表に罫線を引き、金額が 5,000 円以上のセルを赤い太字にします。

:::::platform
::::excel
```vb
salesSheet.Range("A1:F1").Interior.Color = RGB(31, 78, 121)
salesSheet.Range("A1:F1").Font.Color = RGB(255, 255, 255)
salesSheet.Range("A1:F" & lastRow).Borders.LineStyle = xlContinuous

For i = 2 To lastRow
    If salesSheet.Cells(i, 6).Value >= 5000 Then
        salesSheet.Cells(i, 6).Font.Color = RGB(192, 0, 0)
        salesSheet.Cells(i, 6).Font.Bold = True
    End If
Next i
```

`RGB(赤, 緑, 青)` は色を 0〜255 の 3 つの数で表します。
::::
::::sheets
```js
salesSheet.getRange('A1:F1').setBackground('#1f4e79').setFontColor('#ffffff');
salesSheet.getRange('A1:F' + lastRow).setBorder(true, true, true, true, true, true);

for (let i = 2; i <= lastRow; i++) {
  const cell = salesSheet.getRange(i, 6);
  if (cell.getValue() >= 5000) {
    cell.setFontColor('#c00000').setFontWeight('bold');
  }
}
```

色は `#` に続く 6 桁（赤・緑・青を 2 桁ずつ）で表します。
::::
:::::

:::notice[それ、条件付き書式でもできます]
「5,000 円以上を赤くする」だけなら、Excel の **条件付き書式** のほうが向いています。値が変われば自動で色も変わるからです。第 3 章の「関数で十分だった」と同じで、VBA は他の機能でできないところに使うのが基本です。
:::

## デモ 2: 顧客ごとにシートを分ける（`SplitByCustomer`）

ここからが関数ではできない操作です。5 社それぞれの名前のシートを作り、その会社の明細だけを書き写します。

:::::platform
::::excel
```vb
For c = 5 To 9
    customerName = salesSheet.Cells(c, 8).Value
    Set customerSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    customerSheet.Name = customerName
    salesSheet.Range("A1:F1").Copy customerSheet.Range("A1")

    outRow = 2
    For i = 2 To lastRow
        If salesSheet.Cells(i, 2).Value = customerName Then
            salesSheet.Range("A" & i & ":F" & i).Copy customerSheet.Range("A" & outRow)
            outRow = outRow + 1
        End If
    Next i
Next c
```

- `Worksheets.Add(After:=...)` … いちばん後ろに新しいシートを足す
- `.Copy 貼り付け先` … 範囲をコピーして貼り付ける
- `outRow` … 書き写し先の次の行。1 行書くたびに 1 増やす

第 3 章の `For` と `If` が、そのまま使われていることに注目してください。
::::
::::sheets
```js
for (let c = 5; c <= 9; c++) {
  const customerName = salesSheet.getRange(c, 8).getValue();
  const customerSheet = ss.insertSheet(customerName, ss.getSheets().length);
  salesSheet.getRange('A1:F1').copyTo(customerSheet.getRange('A1'));

  let outRow = 2;
  for (let i = 2; i <= lastRow; i++) {
    if (salesSheet.getRange(i, 2).getValue() === customerName) {
      salesSheet.getRange('A' + i + ':F' + i).copyTo(customerSheet.getRange('A' + outRow));
      outRow = outRow + 1;
    }
  }
}
```

- `ss.insertSheet(名前, 位置)` … 新しいシートを足す
- `.copyTo(貼り付け先)` … 範囲をコピーして貼り付ける

第 3 章の `for` と `if` が、そのまま使われていることに注目してください。
::::
:::::

### もう一度実行すると…

このマクロをもう一度実行すると、エラーで止まります。「ひまわり商事」というシートがもう存在していて、同じ名前は付けられないからです。

エラーで止まった行と、そのときの `customerName` を確かめてみてください（第 3 章のデバッグと同じ手順です）。やり直すには、「売上」以外のシートを全部消す `DeleteDemoSheets` を実行します。

:::::platform
::::excel
```vb
Application.DisplayAlerts = False
For Each ws In Worksheets
    If ws.Name <> "売上" Then
        ws.Delete
    End If
Next ws
Application.DisplayAlerts = True
```

- `For Each ws In Worksheets` 〜 `Next ws` … ブックのシートを 1 枚ずつ `ws` に入れてくり返す。番号ではなく「全部について」くり返したいときの書き方です
- `Application.DisplayAlerts = False` … 「シートを削除しますか？」の確認を出さないようにする。終わったら `True` に戻します
::::
::::sheets
```js
for (const sheet of ss.getSheets()) {
  if (sheet.getName() !== '売上') {
    ss.deleteSheet(sheet);
  }
}
```

- `for (const sheet of ss.getSheets())` … シートを 1 枚ずつ `sheet` に入れてくり返す。番号ではなく「全部について」くり返したいときの書き方です
::::
:::::

## デモ 3: ファイルを保存する（`SaveBackup`）

日時入りの名前で、ブックのコピーを保存します。毎月の締めの前にバックアップを取る、といった使い方ができます。

:::::platform
::::excel
```vb
fileName = ThisWorkbook.Path & Application.PathSeparator & _
    "売上_バックアップ_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"
ThisWorkbook.SaveCopyAs fileName
```

- `ThisWorkbook.Path` … このブックが保存されているフォルダ
- `Application.PathSeparator` … フォルダの区切り文字（Windows は `\`、Mac は `/`）
- 行末の ` _` … 長い行を途中で折り返す記号
::::
::::sheets
```js
const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
DriveApp.getFileById(ss.getId()).makeCopy('売上_バックアップ_' + stamp);
```

コピーは Google ドライブに保存されます。ドライブを操作するので、初回は追加の承認を求められます。
::::
:::::

## 帳票づくりにつなげる

ここまでの部品を組み合わせると、月末の請求書づくりがほぼ書けます。

| 請求書づくりの手順 | 使う部品 | 出てきた章 |
|---|---|---|
| 顧客を 1 社ずつ処理する | `For` | 3 |
| 請求書のひな形をコピーする | シートの `.Copy` | 4 |
| 宛名を書く | `Range(...).Value` | 2 |
| その顧客の明細だけを書き写す | `For` と `If` | 3・4 |
| PDF にする | シートの `.ExportAsFixedFormat` | 5 |
| ボタンで実行する | ボタンへのマクロの登録 | 5 |

## どの道具を使うか

| やりたいこと | まず考える道具 |
|---|---|
| 値を計算する・集計する | 関数（SUM、SUMIF など） |
| 条件によって見た目を変える | 条件付き書式 |
| 条件に合う行だけを見る | フィルター、FILTER 関数 |
| シートやファイルを作る・消す・保存する | VBA |
| 毎回同じ手順の操作をまとめて実行する | VBA |

VBA は強力ですが、関数や標準の機能で済むなら、そのほうが誰にでも分かりやすく、壊れにくい仕組みになります。

## 次へ

[テンプレートをコピーして宛名を書く](../../05-invoice-button/01-copy-template/LECTURE.md) から、5 社分の請求書をボタンひとつで作ります。

::codeview[この節のデモのコード]

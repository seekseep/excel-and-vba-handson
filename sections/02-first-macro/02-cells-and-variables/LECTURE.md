---
docs: true
title: セルと変数
---

# セルと変数

![セルから読んで、変数に入れて、計算して、セルに書く](./images/00-thumbnail.svg)

2-1 で作ったマクロの下に続けて書いていきます。セルに書く、セルから読んで計算する。この 2 つができれば、第 3 章の集計に必要な道具はそろいます。

## セルに書き込む

次は、メッセージではなくセルに書いてみます。さっきのマクロの下に続けて書きます。

:::::platform
::::excel
```vb
Sub WriteCells()
    Range("A1").Value = "税抜価格"
    Range("B1").Value = 1200
    Range("A2").Value = "税込価格"
End Sub
```

- `Range("A1")` … A1 のセル
- `.Value` … そのセルの値
- `=` … **右の値を左に入れる**。算数の「等しい」ではありません
::::
::::sheets
```js
function writeCells() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('A1').setValue('税抜価格');
  sheet.getRange('B1').setValue(1200);
  sheet.getRange('A2').setValue('税込価格');
}
```

- `SpreadsheetApp.getActiveSheet()` … いま開いているシート。`const sheet =` で `sheet` という名前を付けておき、以降はそれを使います
- `sheet.getRange('A1')` … そのシートの A1 のセル
- `.setValue(...)` … セルに値を書く
::::
:::::

実行して、シートの A1・B1・A2 に値が入ったことを確かめてください。ブラウザ上で再現したデモでも試せます。

:::::platform
::::excel
「開発」タブの「マクロ」を押し、一覧から `WriteCells` を選んで「実行」を押します。
::::
::::sheets
Apps Script エディタの関数の選択欄で `writeCells` を選び、「▶ 実行」を押します。
::::
:::::

::::sheetdemo[WriteCells を実行すると、A1・B1・A2 に値が入る]{macro="WriteCells" macros="Hello,WriteCells" stage="write-cells"}
A1:B1 = 税抜価格, 1200
A2 = 税込価格
::::

セルは「行と列の番号」でも指せます。A1 は 1 行目の 1 列目、B1 は 1 行目の 2 列目です。

| 番地で指す | 番号で指す |
|---|---|
| `Range("B1")` | `Cells(1, 2)` |
| `sheet.getRange('B1')` | `sheet.getRange(1, 2)` |

番号で指す書き方は、第 3 章で「2 行目から 31 行目まで順番に」処理するときに使います。

## セルから読んで、計算して、書く

B1 の税抜価格を読んで、税込価格（10%）を B2 に書きます。

:::::platform
::::excel
```vb
Sub CalcTaxIncluded()
    Dim price As Long
    Dim taxIncluded As Long

    price = Range("B1").Value
    taxIncluded = Int(price * 1.1)
    Range("B2").Value = taxIncluded
End Sub
```

- `Dim price As Long` … `price` という名前の **変数**（値を入れておく箱）を用意する。`As Long` は「整数を入れる箱」という意味
- `price = Range("B1").Value` … B1 の値を `price` に入れる
- `Int(price * 1.1)` … 1.1 倍して、小数点以下を切り捨てる
::::
::::sheets
```js
function calcTaxIncluded() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const price = sheet.getRange('B1').getValue();
  const taxIncluded = Math.floor(price * 1.1);
  sheet.getRange('B2').setValue(taxIncluded);
}
```

- `const price = ...` … `price` という名前の **変数**（値を入れておく箱）を作って、B1 の値を入れる
- `.getValue()` … セルの値を読む
- `Math.floor(price * 1.1)` … 1.1 倍して、小数点以下を切り捨てる
::::
:::::

B2 に `1320` と出れば成功です。

:::::platform
::::excel
「開発」タブの「マクロ」を押し、一覧から `CalcTaxIncluded` を選んで「実行」を押します。
::::
::::sheets
Apps Script エディタの関数の選択欄で `calcTaxIncluded` を選び、「▶ 実行」を押します。
::::
:::::

::::sheetdemo[CalcTaxIncluded を実行すると、B2 に税込価格が入る]{macro="CalcTaxIncluded" macros="Hello,WriteCells,CalcTaxIncluded"}
before A1:B1 = 税抜価格, 1200
before A2 = 税込価格
B2 = 1320
::::

B1 を `3000` に書き換えてもう一度実行すると、B2 は `3300` になります。

### 型と計算

VBA の変数には、何を入れる箱かを `As` で指定します（**型**）。この教材で使うのは次の 4 つです。

| 型 | 入れるもの | 例 |
|---|---|---|
| `Long` | 整数 | 行番号、個数、金額 |
| `String` | 文字列 | 顧客名、シート名 |
| `Double` | 小数 | 割合 |
| `Date` | 日付 | 発行日 |

計算の記号は `+`（足す）`-`（引く）`*`（掛ける）`/`（割る）です。

GAS では型を書きません。そのかわり、あとで中身を変える変数は `let`、変えない変数は `const` で作ります。

:::notice[変数は必要なときだけ]
Excel ではセルそのものがデータの置き場所です。`Range("B1").Value * 1.1` のように、セルを直接使っても計算できます。変数は「何度も使う値」や「名前を付けたほうが読みやすい値」に絞って使えば十分です。
:::

## 保存する

:::::platform
::::excel
「ファイル」→「名前を付けて保存」で、ファイルの種類を **Excel マクロ有効ブック（\*.xlsm）** にして保存します。`.xlsx` で保存するとマクロが消えるので注意してください。
::::
::::sheets
スプレッドシートもスクリプトも自動で保存されます（エディタは `Ctrl` + `S` で保存）。
::::
:::::

## やってみよう

1. 送料 500 円を足した「支払金額」を A3・B3 に出すマクロを書く
2. `Hello` のメッセージを、B2 の税込価格を表示するように変える（ヒント: VBA は `"税込価格は " & Range("B2").Value & " 円です"`、GAS は `+` で文字列をつなぐ）

::codeview[この節の完成コード]

:::faq

### マクロで書き込んだ結果を `Ctrl` + `Z` で元に戻せません

VBA の実行結果は、元に戻す（`Ctrl` + `Z`）の対象になりません。大事なブックで試すときは、実行の前に保存しておくか、コピーしたブックで試してください。スプレッドシートは「ファイル」→「変更履歴」から前の状態に戻せます。

### 保存しようとしたら「マクロなしのブックに保存できません」と出ました

`.xlsx` のまま保存しようとしています。ここで「はい」を押すと、マクロを消して保存されてしまいます。「いいえ」を押し、ファイルの種類を **Excel マクロ有効ブック（\*.xlsm）** にして保存し直してください。

### `Range("B1").Value` の `.Value` は省略できますか？

省略しても動きます。`Range` の既定のプロパティが `Value` だからです。ただ、第 4 章でセルを変数に入れるようになると、「セルそのもの」なのか「セルの値」なのかが紛らわしくなります。この教材では必ず `.Value` を書きます。

### 整数なのに、なぜ `Integer` ではなく `Long` を使うのですか？

`Integer` には -32,768〜32,767 までしか入らないので、行番号や金額がすぐにあふれて「オーバーフローしました」のエラーになります。今の VBA では `Long` を使っても遅くなったりメモリを余計に使ったりはしないので、整数はすべて `Long` で書きます。

### B2 に `#####` と表示されます

値が入っていないのではなく、列の幅が足りずに数字を表示しきれていない状態です。列見出し（B と C）の境目をダブルクリックすると、幅が中身に合わせて広がります。

### `Range("B1")` と `Cells(1, 2)`、どちらで書けばよいですか？

決まった 1 つのセルを指すなら `Range("B1")` のほうが読みやすく、行や列を変数で動かすなら `Cells(i, 2)` が便利です。範囲を指すときは `Range(Cells(2, 1), Cells(lastRow, 6))` のように組み合わせることもできます。GAS の `getRange` も、`'B1'` と `(1, 2)` の両方の書き方を受け付けます。

:::

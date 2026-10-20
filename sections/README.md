---
docs: true
title: Excel と Google スプレッドシート
---

# Excel と Google スプレッドシート

この教材は **Excel（VBA）** と **Google スプレッドシート（Apps Script、以下 GAS）** の両方に対応しています。説明は Excel を主に書いていますが、同じ課題をスプレッドシートでもまったく同じ流れで進められます。

どちらで進めるかを選んでください。

::platform-choice

選ぶと、手順やコードが分かれる箇所が、すべてのページでその環境のものになります。たとえば次のコードは、選んでいる環境のものだけが表示されます。

:::::platform
::::excel
A1 のセルに「こんにちは」と書くマクロ（VBA）:

```vb
Sub Hello()
    Range("A1").Value = "こんにちは"
End Sub
```

VBE（`Alt` + `F11`）の標準モジュールに書いて、`F5` で実行します。
::::
::::sheets
A1 のセルに「こんにちは」と書くマクロ（GAS）:

```js
function hello() {
  SpreadsheetApp.getActiveSheet().getRange('A1').setValue('こんにちは');
}
```

「拡張機能」→「Apps Script」のエディタに書いて、「実行」を押します。
::::
:::::

あとから変えたいときは、画面右上のナビゲーションバー（テーマ切り替えの左）でも選べます。選んだ環境はブラウザが覚えておくので、次に開いたときもそのままです。

- 会社や学校の PC で Excel のマクロが禁止されている
- Excel を持っていない

という場合は、スプレッドシートで参加してください。VBA と GAS の書き方は [基本的な知識](./00-setup/03-basics/LECTURE.md) にまとめてあります。

## それぞれの準備

- Excel で進める人は [Excel の準備](./00-setup/01-excel/LECTURE.md)
- スプレッドシートで進める人は [スプレッドシートの準備](./00-setup/02-sheets/LECTURE.md)

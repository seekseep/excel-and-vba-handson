# シートとファイルを操作する

本文は [LECTURE.md](./LECTURE.md)。

```text
example/
├── data/sales.xlsx
├── gas/ObjectDemo.gs
├── vba/ObjectDemo.bas
```

`example/` が配布 ZIP と、本文末尾の完成コードの元です。

## デモの順番

1. `ShowObjects` — オブジェクトの入れ子をイミディエイト ウィンドウで見せる
2. `DecorateSales` — 書式（条件付き書式で十分、という話につなげる）
3. `SplitByCustomer` — 2 回実行してエラーを出し、デバッグの復習をする
4. `DeleteDemoSheets` — `For Each` の紹介
5. `SaveBackup` — ファイル保存

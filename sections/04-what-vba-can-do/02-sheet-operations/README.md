# シートとファイルを操作する

本文は [LECTURE.md](./LECTURE.md)。

```text
├── book/sales.xlsm
├── book/sales.xlsx
├── example/gas/ObjectDemo.gs
└── example/vba/ObjectDemo.bas
```

`example/` は本文末尾の完成コードの元です。`book/` は練習用ブックで、`tools/build_workbooks.py` が作ります（`.xlsm` は Excel で配る節を終えた状態、`.xlsx` は表計算デモが描く節を始める状態）。

## デモの順番

1. `ShowObjects` — オブジェクトの入れ子をイミディエイト ウィンドウで見せる
2. `DecorateSales` — 書式（条件付き書式で十分、という話につなげる）
3. `SplitByCustomer` — 2 回実行してエラーを出し、デバッグの復習をする
4. `DeleteDemoSheets` — `For Each` の紹介
5. `SaveBackup` — ファイル保存

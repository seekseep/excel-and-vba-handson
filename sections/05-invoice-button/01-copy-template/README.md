# テンプレートをコピーして宛名を書く

本文は [LECTURE.md](./LECTURE.md)。

```text
├── book/invoice.xlsm
├── book/invoice.xlsx
├── example/gas/Invoice.gs
└── example/vba/Invoice.bas
```

`example/` は本文末尾の完成コードの元です。`book/` は練習用ブックで、`tools/build_workbooks.py` が作ります（`.xlsm` は Excel で配る節を終えた状態、`.xlsx` は表計算デモが描く節を始める状態）。

## テンプレートのセル番地

請求書テンプレートのセル番地（A5 宛名 / B9 請求金額 / 12〜31 行目 明細 / E33 合計）は、`tools/build_workbooks.py` とコードの両方に書いてあります。レイアウトを変えるときは両方を直してください。

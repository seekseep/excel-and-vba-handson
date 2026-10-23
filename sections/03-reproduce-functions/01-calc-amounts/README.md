# 金額を計算する

本文は [LECTURE.md](./LECTURE.md)。

```text
├── book/sales.xlsm
├── book/sales.xlsx
├── example/gas/SalesSummary.gs
├── example/vba/SalesSummary.bas
├── stages/for-31/gas/SalesSummary.gs
├── stages/for-31/vba/SalesSummary.bas
├── stages/one-row/gas/SalesSummary.gs
└── stages/one-row/vba/SalesSummary.bas
```

`example/` は本文末尾の完成コードの元です。`stages/` は本文の途中のデモに出すその時点のコードです（`one-row` は 1 行分、`for-31` は 31 行目までの For）。`book/` は練習用ブックで、`tools/build_workbooks.py` が作ります（`.xlsm` は Excel で配る節を終えた状態、`.xlsx` は表計算デモが描く節を始める状態）。

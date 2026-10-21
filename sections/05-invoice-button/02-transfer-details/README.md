# 明細を書き写す

本文は [LECTURE.md](./LECTURE.md)。

```text
├── book/invoice.xlsm
├── example/gas/Invoice.gs
├── example/vba/Invoice.bas
├── stages/one-customer/   ひまわり商事 1 社分の明細まで（本文 1 つ目のデモが表示する）
└── stages/all-customers/  5 社分のくり返しまで。DeleteInvoices の手前（本文 2 つ目のデモが表示する）
```

`example/` は本文末尾の完成コードの元です。`book/` は練習用ブックで、`tools/build_workbooks.py` が作ります（`.xlsm` は Excel で配る節を終えた状態、`.xlsx` は表計算デモが描く節を始める状態）。

# テンプレートをコピーして宛名を書く

本文は [LECTURE.md](./LECTURE.md)。

```text
example/
├── data/invoice.xlsx
├── gas/Invoice.gs
├── vba/Invoice.bas
```

`example/` が配布 ZIP と、本文末尾の完成コードの元です。

## テンプレートのセル番地

請求書テンプレートのセル番地（A5 宛名 / B9 請求金額 / 12〜31 行目 明細 / E33 合計）は、`tools/build_workbooks.py` とコードの両方に書いてあります。レイアウトを変えるときは両方を直してください。

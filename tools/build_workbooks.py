#!/usr/bin/env python3
"""
各レクチャーの example/data/ に置く練習用ブックを生成する。

  - .xlsx … マクロなし。章の最初の節を始めるときと、スプレッドシートで使う（アップロードして変換する）
  - .xlsm … .xlsx と同じ中身に、その節を終えた状態のマクロ（example/vba/*.bas）を入れたもの。
           Excel 用の ZIP にはこちらだけを入れる（build-downloads.mjs）
.xlsm のマクロの部分（vbaProject.bin）は tools/vba_project.py が仕様どおりに組み立てる（Excel は不要）。

使い方（保守者のみ。受講者は実行しない）:
    pip install openpyxl
    python3 tools/build_workbooks.py

売上データは乱数の種を固定しているので、何度実行しても同じ中身になる。
"""

import random
import zipfile
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from vba_project import build_vba_project

ROOT = Path(__file__).resolve().parent.parent
SECTIONS = ROOT / "sections"

SHOP = {
    "name": "そよかぜ文具店",
    "zip": "〒150-0000",
    "address": "東京都渋谷区そよかぜ町 1-2-3",
    "tel": "TEL 03-0000-0000",
}

CUSTOMERS = [
    # 顧客名, 郵便番号, 住所, ご担当者
    ("ひまわり商事", "〒104-0000", "東京都中央区ひまわり 2-4-6", "総務部 山田 様"),
    ("あおば設計事務所", "〒160-0000", "東京都新宿区あおば台 3-1-8", "事務担当 佐藤 様"),
    ("みなと物産", "〒231-0000", "神奈川県横浜市中区みなと 5-5-1", "経理部 鈴木 様"),
    ("さくら学習塾", "〒180-0000", "東京都武蔵野市さくら通り 1-9", "教室長 高橋 様"),
    ("こもれび珈琲", "〒154-0000", "東京都世田谷区こもれび 7-2-10", "店長 田中 様"),
]

PRODUCTS = [
    # 商品, 単価
    ("ボールペン", 120),
    ("ノート", 180),
    ("クリアファイル", 60),
    ("付箋", 250),
    ("コピー用紙", 550),
    ("ホッチキス", 680),
    ("修正テープ", 320),
]

SALES_ROWS = 30

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
BOLD = Font(bold=True)
THIN = Side(style="thin", color="808080")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
YEN = '#,##0'
DATE = 'yyyy/mm/dd'


def make_sales():
    """2026 年 10 月の売上明細 30 行。日付順に並べる。"""
    rng = random.Random(20261001)
    rows = []
    for _ in range(SALES_ROWS):
        day = rng.randint(1, 30)
        customer = rng.choice(CUSTOMERS)[0]
        product, price = rng.choice(PRODUCTS)
        quantity = rng.choice([1, 2, 3, 5, 10, 12, 20])
        rows.append((date(2026, 10, day), customer, product, price, quantity))
    rows.sort(key=lambda r: (r[0], r[1]))
    return rows


def write_sales_sheet(ws, sales, with_amount):
    headers = ["日付", "顧客", "商品", "単価", "数量", "金額"]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.border = BOX
    for r, (d, customer, product, price, quantity) in enumerate(sales, start=2):
        values = [d, customer, product, price, quantity, price * quantity if with_amount else None]
        for col, v in enumerate(values, start=1):
            cell = ws.cell(row=r, column=col, value=v)
            cell.border = BOX
        ws.cell(row=r, column=1).number_format = DATE
        ws.cell(row=r, column=4).number_format = YEN
        ws.cell(row=r, column=6).number_format = YEN
    widths = {"A": 12, "B": 18, "C": 16, "D": 9, "E": 7, "F": 10}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"


def write_summary_area(ws):
    """3 章の集計欄。I 列に関数、J 列に VBA の結果を並べて比べる。"""
    def header(cell_ref, value):
        cell = ws[cell_ref]
        cell.value = value
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.border = BOX

    header("H1", "項目")
    header("I1", "関数")
    header("J1", "VBA")
    ws["H2"] = "売上合計"
    ws["H2"].border = BOX
    for ref in ("I2", "J2"):
        ws[ref].border = BOX
        ws[ref].number_format = YEN
    ws["I2"].fill = INPUT_FILL

    header("H4", "顧客")
    header("I4", "関数")
    header("J4", "VBA")
    for i, (name, *_rest) in enumerate(CUSTOMERS, start=5):
        ws.cell(row=i, column=8, value=name).border = BOX
        for col in (9, 10):
            cell = ws.cell(row=i, column=col)
            cell.border = BOX
            cell.number_format = YEN
        ws.cell(row=i, column=9).fill = INPUT_FILL

    ws.column_dimensions["G"].width = 3
    ws.column_dimensions["H"].width = 18
    ws.column_dimensions["I"].width = 12
    ws.column_dimensions["J"].width = 12


def write_customer_sheet(ws):
    headers = ["顧客名", "郵便番号", "住所", "ご担当者"]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.border = BOX
    for r, row in enumerate(CUSTOMERS, start=2):
        for col, v in enumerate(row, start=1):
            ws.cell(row=r, column=col, value=v).border = BOX
    for col, w in {"A": 18, "B": 12, "C": 32, "D": 18}.items():
        ws.column_dimensions[col].width = w


def write_invoice_template(ws):
    """
    請求書のひな形。マクロはこのシートをコピーして値を埋める。
    セル番地は VBA / GAS のコードと対応しているので、動かすときは両方を直すこと。
      A5 宛名 / A6 住所 / A7 ご担当者 / E3 発行日 / B9 ご請求金額
      12〜31 行目 明細（A 日付, B 商品, C 単価, D 数量, E 金額） / E33 合計
    """
    for col, w in {"A": 14, "B": 22, "C": 12, "D": 10, "E": 14}.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells("A1:E1")
    ws["A1"] = "請 求 書"
    ws["A1"].font = Font(size=20, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")

    ws["D3"] = "発行日"
    ws["D3"].alignment = Alignment(horizontal="right")
    ws["E3"].number_format = DATE

    ws["A5"] = "（宛名）"
    ws["A5"].font = Font(size=14, bold=True)
    ws["A6"] = "（住所）"
    ws["A7"] = "（ご担当者）"

    ws["D5"] = SHOP["name"]
    ws["D5"].font = BOLD
    ws["D6"] = SHOP["zip"]
    ws["D7"] = SHOP["address"]
    ws["D8"] = SHOP["tel"]

    ws["A9"] = "ご請求金額"
    ws["A9"].font = BOLD
    ws["B9"].number_format = '"¥"#,##0'
    ws["B9"].font = Font(size=14, bold=True)
    ws["B9"].border = Border(bottom=Side(style="medium"))

    headers = ["日付", "商品", "単価", "数量", "金額"]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=11, column=col, value=h)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center")
    for r in range(12, 32):
        for col in range(1, 6):
            ws.cell(row=r, column=col).border = BOX
        ws.cell(row=r, column=1).number_format = DATE
        ws.cell(row=r, column=3).number_format = YEN
        ws.cell(row=r, column=5).number_format = YEN

    ws["D33"] = "合計"
    ws["D33"].font = BOLD
    ws["D33"].border = BOX
    ws["E33"].number_format = YEN
    ws["E33"].font = BOLD
    ws["E33"].border = BOX

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = "A1:E33"


def save(wb, path, end_state=None):
    """
    .xlsx（節を始める状態。本文のデモもこれを描く）を保存し、マクロ入りの .xlsm も作る。
    end_state(wb) を渡すと、.xlsm はその節のマクロを実行し終えた後の値まで入れた状態にする
    （次の節を、ダウンロードした .xlsm からそのまま始められるように）。
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    print(f"[build_workbooks] {path.relative_to(ROOT)}")
    if end_state:
        end_state(wb)
    save_xlsm(wb, path)


def save_xlsm(wb, xlsx_path):
    """
    その節の完成コード（example/vba/*.bas）があれば、同じ中身のマクロ入りブック（.xlsm）も作る。
    openpyxl はマクロを書けないので、.xlsx を保存してから zip の中身を差し替える:
      - xl/vbaProject.bin を足す（tools/vba_project.py で組み立てたもの）
      - ブックの種類を「マクロ有効」に変え、vbaProject.bin へのリレーションを足す
      - ブックとシートにコード名（ThisWorkbook / Sheet1 …）を付ける（VBA の部品と対応させるため）
    """
    vba_dir = xlsx_path.parent.parent / "vba"
    bas_files = sorted(vba_dir.glob("*.bas"))
    if not bas_files:
        return
    wb.code_name = "ThisWorkbook"
    sheet_names = []
    for i, ws in enumerate(wb.worksheets, start=1):
        ws.sheet_properties.codeName = f"Sheet{i}"
        sheet_names.append(f"Sheet{i}")
    xlsm_path = xlsx_path.with_suffix(".xlsm")
    tmp = xlsx_path.with_suffix(".tmp.xlsx")
    wb.save(tmp)
    vba = build_vba_project([(p.stem, p.read_text()) for p in bas_files], sheet_names)
    with zipfile.ZipFile(tmp) as zin, zipfile.ZipFile(xlsm_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                text = data.decode()
                text = text.replace(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                )
                text = text.replace(
                    '<Default Extension="xml"',
                    '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="xml"',
                    1,
                )
                data = text.encode()
            elif item.filename == "xl/_rels/workbook.xml.rels":
                text = data.decode().replace(
                    "</Relationships>",
                    '<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
                )
                data = text.encode()
            zout.writestr(item, data)
        zout.writestr("xl/vbaProject.bin", vba)
    tmp.unlink()
    print(f"[build_workbooks] {xlsm_path.relative_to(ROOT)}（マクロ入り）")


def main():
    sales = make_sales()

    # 2 章: 空白のブックから始める。終えた状態として、空のシートにマクロを入れたブックを作る
    for lecture in ("01-hello-macro", "02-cells-and-variables"):
        wb = Workbook()
        wb.active.title = "Sheet1"
        path = SECTIONS / "02-first-macro" / lecture / "example" / "data" / "first-macro.xlsx"
        path.parent.mkdir(parents=True, exist_ok=True)
        save_xlsm(wb, path)

    # 3 章: 金額が空の売上明細 + 集計欄（受講者が関数と VBA で埋める）。.xlsx はどの節も同じ
    # .xlsm は、その節のマクロを実行し終えた後の値（金額・合計・顧客別の合計）まで入れる
    totals = {}
    for _, customer, _, price, quantity in sales:
        totals[customer] = totals.get(customer, 0) + price * quantity

    def fill_amounts(ws):
        for r, (_, _, _, price, quantity) in enumerate(sales, start=2):
            ws.cell(row=r, column=6, value=price * quantity)

    def end_31(wb):
        fill_amounts(wb["売上"])

    def end_32(wb):
        ws = wb["売上"]
        fill_amounts(ws)
        ws["J2"] = sum(totals.values())
        ws["J5"] = totals[CUSTOMERS[0][0]]

    def end_34(wb):
        end_32(wb)
        for i, (name, *_rest) in enumerate(CUSTOMERS, start=5):
            wb["売上"].cell(row=i, column=10, value=totals[name])

    for lecture, end_state in (
        ("01-calc-amounts", end_31),
        ("02-sum-and-sumif", end_32),
        ("03-debugger-intro", end_32),
        ("04-debugging", end_34),
    ):
        wb = Workbook()
        ws = wb.active
        ws.title = "売上"
        write_sales_sheet(ws, sales, with_amount=False)
        write_summary_area(ws)
        save(wb, SECTIONS / "03-reproduce-functions" / lecture / "example" / "data" / "sales.xlsx", end_state)

    # 4 章: 同じ表を金額まで埋めた状態（デモは集計済みの表から始める）
    for lecture in ("01-object-model", "02-sheet-operations"):
        wb = Workbook()
        ws = wb.active
        ws.title = "売上"
        write_sales_sheet(ws, sales, with_amount=True)
        write_summary_area(ws)
        save(wb, SECTIONS / "04-what-vba-can-do" / lecture / "example" / "data" / "sales.xlsx")

    # 5 章: 金額まで埋まった売上明細 + 顧客 + 請求書テンプレート
    for lecture in ("01-copy-template", "02-transfer-details", "03-pdf-and-button"):
        wb = Workbook()
        ws = wb.active
        ws.title = "売上"
        write_sales_sheet(ws, sales, with_amount=True)
        write_customer_sheet(wb.create_sheet("顧客"))
        write_invoice_template(wb.create_sheet("請求書テンプレート"))
        save(wb, SECTIONS / "05-invoice-button" / lecture / "example" / "data" / "invoice.xlsx")


if __name__ == "__main__":
    main()

Option Explicit

' 5-1 テンプレートをコピーして宛名を書く
' data/invoice.xlsx を .xlsm で保存してから実行する。
' 実行するたびに「請求書_ひまわり商事」ができるので、2 回目の前にそのシートを手で消しておく。

' 顧客シートの 2 行目（ひまわり商事）の請求書シートを 1 枚作る
Sub CreateInvoices()
    Dim salesSheet As Worksheet
    Dim customerSheet As Worksheet
    Dim invoiceSheet As Worksheet
    Dim c As Long
    Dim customerName As String

    Set salesSheet = Worksheets("売上")
    Set customerSheet = Worksheets("顧客")

    c = 2
    customerName = customerSheet.Cells(c, 1).Value

    ' テンプレートを末尾にコピーして名前を付ける
    Worksheets("請求書テンプレート").Copy After:=Worksheets(Worksheets.Count)
    Set invoiceSheet = Worksheets(Worksheets.Count)
    invoiceSheet.Name = "請求書_" & customerName

    ' 宛名と発行日
    invoiceSheet.Range("A5").Value = customerName & " 御中"
    invoiceSheet.Range("A6").Value = customerSheet.Cells(c, 2).Value & " " & customerSheet.Cells(c, 3).Value
    invoiceSheet.Range("A7").Value = customerSheet.Cells(c, 4).Value
    invoiceSheet.Range("E3").Value = Date
End Sub

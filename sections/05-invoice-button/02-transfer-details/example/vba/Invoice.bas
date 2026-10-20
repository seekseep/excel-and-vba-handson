Option Explicit

' 5-2 明細を書き写し、全顧客分を作る
' data/invoice.xlsx を .xlsm で保存してから実行する。
' 請求書テンプレートのセル番地（A5 宛名 / B9 請求金額 / 12 行目から明細 / E33 合計）に合わせて書いている。

' 顧客ごとに請求書シートを作る
Sub CreateInvoices()
    Dim salesSheet As Worksheet
    Dim customerSheet As Worksheet
    Dim invoiceSheet As Worksheet
    Dim lastSalesRow As Long
    Dim lastCustomerRow As Long
    Dim c As Long
    Dim i As Long
    Dim lineRow As Long
    Dim total As Long
    Dim customerName As String

    ' 何度実行しても同じ結果になるよう、前回の請求書シートを消してから作る
    DeleteInvoices

    Set salesSheet = Worksheets("売上")
    Set customerSheet = Worksheets("顧客")
    lastSalesRow = salesSheet.Cells(salesSheet.Rows.Count, 1).End(xlUp).Row
    lastCustomerRow = customerSheet.Cells(customerSheet.Rows.Count, 1).End(xlUp).Row

    For c = 2 To lastCustomerRow
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

        ' 売上からこの顧客の行だけを明細に転記する
        lineRow = 12
        total = 0
        For i = 2 To lastSalesRow
            If salesSheet.Cells(i, 2).Value = customerName Then
                invoiceSheet.Cells(lineRow, 1).Value = salesSheet.Cells(i, 1).Value
                invoiceSheet.Cells(lineRow, 2).Value = salesSheet.Cells(i, 3).Value
                invoiceSheet.Cells(lineRow, 3).Value = salesSheet.Cells(i, 4).Value
                invoiceSheet.Cells(lineRow, 4).Value = salesSheet.Cells(i, 5).Value
                invoiceSheet.Cells(lineRow, 5).Value = salesSheet.Cells(i, 6).Value
                total = total + salesSheet.Cells(i, 6).Value
                lineRow = lineRow + 1
            End If
        Next i
        invoiceSheet.Range("E33").Value = total
        invoiceSheet.Range("B9").Value = total
    Next c

    customerSheet.Activate
    MsgBox (lastCustomerRow - 1) & " 件の請求書を作りました"
End Sub

' 「請求書_」で始まるシートを消す
Sub DeleteInvoices()
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If Left(ws.Name, 4) = "請求書_" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub

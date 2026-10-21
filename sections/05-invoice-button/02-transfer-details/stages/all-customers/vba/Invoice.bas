Option Explicit

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

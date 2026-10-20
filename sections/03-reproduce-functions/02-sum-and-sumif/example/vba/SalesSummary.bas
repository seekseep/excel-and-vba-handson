Option Explicit

Sub CalcAmounts()
    Dim lastRow As Long
    Dim i As Long

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        Cells(i, 6).Value = Cells(i, 4).Value * Cells(i, 5).Value
    Next i
End Sub

Sub SumTotal()
    Dim lastRow As Long
    Dim i As Long
    Dim total As Long

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    total = 0
    For i = 2 To lastRow
        total = total + Cells(i, 6).Value
    Next i
    Range("J2").Value = total
End Sub

Sub SumByCustomer()
    Dim lastRow As Long
    Dim i As Long
    Dim customerName As String
    Dim total As Long

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    customerName = Range("H5").Value
    total = 0
    For i = 2 To lastRow
        If Cells(i, 2).Value = customerName Then
            total = total + Cells(i, 6).Value
        End If
    Next i
    Range("J5").Value = total
End Sub

Option Explicit

Sub SumTotalWithBug()
    Dim lastRow As Long
    Dim i As Long
    Dim total As Long

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    total = 0
    For i = 3 To lastRow
        total = total + Cells(i, 6).Value
    Next i
    Range("J2").Value = total
End Sub

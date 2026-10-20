Option Explicit

Sub CalcAmounts()
    Dim lastRow As Long
    Dim i As Long

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        Cells(i, 6).Value = Cells(i, 4).Value * Cells(i, 5).Value
    Next i
End Sub

Option Explicit

Sub CalcAmounts()
    Dim i As Long

    For i = 2 To 31
        Cells(i, 6).Value = Cells(i, 4).Value * Cells(i, 5).Value
    Next i
End Sub

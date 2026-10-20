Option Explicit

' 3-1 金額を計算する
' data/sales.xlsx の「売上」シートを開いた状態で実行する。

' 金額 = 単価 × 数量 を F 列の全行に書く（関数なら F2 に =D2*E2 を入れて下までコピー）
Sub CalcAmounts()
    Dim lastRow As Long
    Dim i As Long

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        Cells(i, 6).Value = Cells(i, 4).Value * Cells(i, 5).Value
    Next i
End Sub

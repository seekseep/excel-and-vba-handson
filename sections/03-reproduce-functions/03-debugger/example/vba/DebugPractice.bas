Option Explicit

' 3-3 デバッグの練習
' このマクロは売上合計を J2 に書くが、関数（I2）の結果と合わない。
' ブレークポイントとステップ実行で、どこがおかしいかを探そう。

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

Option Explicit

' 3-3 デバッガーで原因を探す
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

' 売上合計を J2 に書く（関数なら =SUM(F2:F31)）
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

' H5 の顧客の合計を J5 に書く（関数なら =SUMIF(B2:B31, H5, F2:F31)）
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

' チャレンジ: H5〜H9 の全顧客の合計を J5〜J9 に書く（For の中に For）
Sub SumAllCustomers()
    Dim lastRow As Long
    Dim c As Long
    Dim i As Long
    Dim total As Long

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For c = 5 To 9
        total = 0
        For i = 2 To lastRow
            If Cells(i, 2).Value = Cells(c, 8).Value Then
                total = total + Cells(i, 6).Value
            End If
        Next i
        Cells(c, 10).Value = total
    Next c
End Sub

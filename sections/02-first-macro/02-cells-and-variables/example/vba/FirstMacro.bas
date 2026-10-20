Option Explicit

Sub Hello()
    MsgBox "こんにちは、VBA！"
End Sub

Sub WriteCells()
    Range("A1").Value = "税抜価格"
    Range("B1").Value = 1200
    Range("A2").Value = "税込価格"
End Sub

Sub CalcTaxIncluded()
    Dim price As Long
    Dim taxIncluded As Long

    price = Range("B1").Value
    taxIncluded = Int(price * 1.1)
    Range("B2").Value = taxIncluded
End Sub

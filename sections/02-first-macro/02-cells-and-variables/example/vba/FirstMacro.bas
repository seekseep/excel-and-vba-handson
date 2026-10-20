Option Explicit

' 2-2 セルと変数
' 2-1 のモジュールに続けて書き、上から順に実行する。

' メッセージを出すだけのマクロ
Sub Hello()
    MsgBox "こんにちは、VBA！"
End Sub

' セルに書き込む
Sub WriteCells()
    Range("A1").Value = "税抜価格"
    Range("B1").Value = 1200
    Range("A2").Value = "税込価格"
End Sub

' セルから読んで、計算して、セルに書く
Sub CalcTaxIncluded()
    Dim price As Long
    Dim taxIncluded As Long

    price = Range("B1").Value
    taxIncluded = Int(price * 1.1)
    Range("B2").Value = taxIncluded
End Sub

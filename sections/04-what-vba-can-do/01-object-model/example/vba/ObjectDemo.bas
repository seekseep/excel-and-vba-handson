Option Explicit

' 4-1 Excel は部品の入れ子
' data/sales.xlsx を .xlsm で保存してから実行し、結果をイミディエイトウィンドウ（Ctrl + G）で見る。

' ブック → シート → セル は入れ子になっている
Sub ShowObjects()
    Debug.Print "ブックの名前: " & ThisWorkbook.Name
    Debug.Print "シートの数: " & ThisWorkbook.Worksheets.Count
    Debug.Print "1 枚目のシート: " & ThisWorkbook.Worksheets(1).Name
    Debug.Print "売上!B2 の値: " & ThisWorkbook.Worksheets("売上").Range("B2").Value
End Sub

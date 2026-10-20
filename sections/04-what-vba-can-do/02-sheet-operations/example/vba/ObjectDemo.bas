Option Explicit

' 4-2 シートとファイルを操作する（講師デモ）
' data/sales.xlsx を .xlsm で保存してから、上から順に実行する。

' ブック → シート → セル は入れ子になっている
Sub ShowObjects()
    Debug.Print "ブックの名前: " & ThisWorkbook.Name
    Debug.Print "シートの数: " & ThisWorkbook.Worksheets.Count
    Debug.Print "1 枚目のシート: " & ThisWorkbook.Worksheets(1).Name
    Debug.Print "売上!B2 の値: " & ThisWorkbook.Worksheets("売上").Range("B2").Value
End Sub

' 表を飾る（色・罫線・条件による文字色）
Sub DecorateSales()
    Dim salesSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set salesSheet = Worksheets("売上")
    lastRow = salesSheet.Cells(salesSheet.Rows.Count, 1).End(xlUp).Row

    salesSheet.Range("A1:F1").Interior.Color = RGB(31, 78, 121)
    salesSheet.Range("A1:F1").Font.Color = RGB(255, 255, 255)
    salesSheet.Range("A1:F" & lastRow).Borders.LineStyle = xlContinuous

    For i = 2 To lastRow
        If salesSheet.Cells(i, 6).Value >= 5000 Then
            salesSheet.Cells(i, 6).Font.Color = RGB(192, 0, 0)
            salesSheet.Cells(i, 6).Font.Bold = True
        End If
    Next i
End Sub

' 顧客ごとにシートを作って明細を振り分ける（関数ではシートを増やせない）
Sub SplitByCustomer()
    Dim salesSheet As Worksheet
    Dim customerSheet As Worksheet
    Dim lastRow As Long
    Dim c As Long
    Dim i As Long
    Dim outRow As Long
    Dim customerName As String

    Set salesSheet = Worksheets("売上")
    lastRow = salesSheet.Cells(salesSheet.Rows.Count, 1).End(xlUp).Row

    For c = 5 To 9
        customerName = salesSheet.Cells(c, 8).Value
        Set customerSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        customerSheet.Name = customerName
        salesSheet.Range("A1:F1").Copy customerSheet.Range("A1")

        outRow = 2
        For i = 2 To lastRow
            If salesSheet.Cells(i, 2).Value = customerName Then
                salesSheet.Range("A" & i & ":F" & i).Copy customerSheet.Range("A" & outRow)
                outRow = outRow + 1
            End If
        Next i
        customerSheet.Columns("A:F").AutoFit
    Next c
    salesSheet.Activate
End Sub

' 「売上」以外のシートを消して、デモをやり直せるようにする
Sub DeleteDemoSheets()
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If ws.Name <> "売上" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub

' 日時入りの名前でブックのコピーを保存する
Sub SaveBackup()
    Dim fileName As String

    fileName = ThisWorkbook.Path & Application.PathSeparator & _
        "売上_バックアップ_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"
    ThisWorkbook.SaveCopyAs fileName
    MsgBox fileName & " に保存しました"
End Sub

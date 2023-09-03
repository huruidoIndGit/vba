Sub ImportCSVData()
    Dim csvFilePath As String
    Dim workFilePath As String
    Dim rowNum As Long
    Dim itemName As String
    Dim ws As Worksheet
    
    ' CSVファイルのパスと作業用ファイルのパスを設定
    csvFilePath = "C:\パス\ファイル.csv" ' CSVファイルのパス
    workFilePath = "C:\パス\作業用ファイル.xlsx" ' 作業用ファイルのパス
    
    ' 作業用ファイルを開く
    Workbooks.Open workFilePath
    Set ws = ActiveSheet ' 作業用ファイルのアクティブシートを設定
    
    ' CSVファイルを開く
    With Workbooks.Open(csvFilePath)
        ' CSVファイルのデータを読み取り
        rowNum = 2 ' 1行目は表題なので2行目から始める
        Do Until .Sheets(1).Cells(rowNum, 1).Value = ""
            itemName = .Sheets(1).Cells(rowNum, 2).Value
            ' 項目名が作業用ファイルに存在するかチェック
            Dim foundCell As Range
            Set foundCell = ws.Range("B:B").Find(itemName, LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not foundCell Is Nothing Then
                ' 項目名が既に存在する場合、個数を増やす
                ws.Cells(foundCell.Row, 3).Value = ws.Cells(foundCell.Row, 3).Value + 1
            Else
                ' 項目名が存在しない場合、新しい行を追加
                Dim lastRow As Long
                lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
                ws.Cells(lastRow + 1, 1).Value = lastRow ' 連番
                ws.Cells(lastRow + 1, 2).Value = itemName ' 項目名
                ws.Cells(lastRow + 1, 3).Value = 1 ' 個数
            End If
            rowNum = rowNum + 1
        Loop
        .Close SaveChanges:=False ' CSVファイルを閉じる
    End With
    
    ' 作業用ファイルを保存
    ActiveWorkbook.Save
    
    ' 作業用ファイルを閉じる
    ActiveWorkbook.Close SaveChanges:=True
End Sub

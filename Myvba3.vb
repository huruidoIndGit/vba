Sub FilterAndExtractData()
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceTable As ListObject
    Dim targetTable As ListObject
    Dim filterColumn As Range
    Dim filterCriteria As String
    Dim newRow As ListRow
    Dim sourceRow As ListRow
    Dim headerRange As Range
    Dim dataRange As Range
    Dim filterEnabled As Boolean
    
    ' ソースシートを指定
    Set sourceSheet = ThisWorkbook.Sheets("元のシートの名前")
    
    ' ソーステーブルを指定
    Set sourceTable = sourceSheet.ListObjects("元のテーブルの名前")
    
    ' フィルタする列を指定 (例: 大項目列)
    Set filterColumn = sourceTable.ListColumns("日付").DataBodyRange
    
    ' フィルタリング条件を指定 (例: "条件A")
    filterCriteria = "2023/1/2"
    
    ' フィルタが適用されているかチェック
    filterEnabled = False
    If sourceSheet.AutoFilterMode Then
        filterEnabled = sourceSheet.AutoFilter.Filters(filterColumn.Column).On
    End If
    
    ' フィルタをクリア
    sourceSheet.AutoFilterMode = False
    
    ' フィルタを適用して条件に合致する行だけを表示
    sourceTable.Range.AutoFilter Field:=filterColumn.Column, Criteria1:=filterCriteria
    
    ' 抽出されたデータをコピーしてターゲットシートに貼り付け
    Set targetSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    Set headerRange = sourceTable.HeaderRowRange
    Set dataRange = sourceTable.DataBodyRange
    
    dataRange.SpecialCells(xlCellTypeVisible).Copy targetSheet.Range("A2")
    headerRange.Copy targetSheet.Range("A1")
    
    ' フィルタをクリア
    sourceSheet.AutoFilterMode = False
    
    ' フィルタを再適用
    If filterEnabled Then
        sourceTable.Range.AutoFilter Field:=filterColumn.Column, Criteria1:=filterCriteria
    End If
    
    ' 新しいテーブルを作成
    Set targetTable = targetSheet.ListObjects.Add(xlSrcRange, targetSheet.UsedRange, , xlYes)

    
    ' メッセージを表示
    MsgBox "データを正常に抽出しました。", vbInformation
End Sub


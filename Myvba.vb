'MainMatrixの文字列を変更するコード
Sub AccessCellsAndOperate()
    Dim ws As Worksheet
    Dim tempRow As Integer
    Dim tempColumn As Variant
    Dim val As Variant  ' セルの値を格納する変数
    Dim bookName
    Dim tempStr
    
    Dim pattern As String
    
    Dim first_Num
    Dim second_Num As Integer  '2つめの数値。1～31の日付。
    Dim third_Num As Integer '3つめの数値。だいたい1～14の番号。
    
    'mainマトリクスの操作対象の列.固定のはず.
    Const columnNumFirstTarget = 10
    Const columnNumLastTarget = 40
    
'=======================実行前に変更する変数===============
'thの番号
    Dim targetNumber As Integer
    targetNumber = 100
'mainマトリクスの操作対象の最初の行番号
    Dim rowNumFirstTarget As Integer
    rowNumFirstTarget = 5
'mainマトリクスの操作対象の最後の行番号
    Dim rowNumLastTarget As Integer
    rowNumLastTarget = 18
'==========================================================
    
    Set ws = ThisWorkbook.Worksheets("Sheet1")     '※Sheet名は必ず「Sheet1」にすること
    bookName = ActiveWorkbook.Name
    '念のため作業対象Book確認
    MsgBox "作業対象Book:  " & ActiveWorkbook.Name
    ' Sheet1 を末尾の位置にコピーする
    Worksheets("Sheet1").Copy After:=Worksheets(Worksheets.Count)
    ws.Activate
    
    
    ' 正規表現パターン
    pattern = "\b\d{1,2}_\d{1,2}_\d{1,2}\b"
    
    ' 正規表現オブジェクトを作成
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = pattern
   
    
    ' mainのセル範囲に順にアクセスして操作
    For tempRow = rowNumFirstTarget To rowNumLastTarget                 '実行前に要変更
        For tempColumn = columnNumFirstTarget To columnNumLastTarget    '固定のはず
            
            '置き換えたい日にちとNumberを取得
            second_Num = ws.Cells(rowNumFirstTarget - 1, tempColumn)
            third_Num = ws.Cells(tempRow, columnNumFirstTarget - 1)
            '変更後のX_X_Xの部分を作成
            Dim newX_X_X As String
            newX_X_X = targetNumber & "_" & second_Num & "_" & third_Num
            
            val = ws.Cells(tempRow, tempColumn).value ' セルの内容を取得
           
            
            ' 正規表現でマッチした部分を"newX_X_X"に置換
            tempStr = regex.Replace(val, newX_X_X)
            ws.Cells(tempRow, tempColumn).value = tempStr
            
        Next tempColumn
    Next tempRow   
    
End Sub



'ランダム文字列と1_2_3を入力するコード

Sub FillCellsWithRandomStrings()
    Dim ws As Worksheet
    Dim rowNum As Long, colNum As Long
    Dim lastRow As Long, lastCol As Long
    Dim charSet As String
    Dim randomString As String
    Dim i As Integer
    
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' シート名を適宜変更
    charSet = "abcdefghijklmnopqrstuvwxyz" ' 生成する文字のセット
    
    
    For rowNum = 5 To 18
        For colNum = 10 To 40
        
            '置き換えたい日にちとNumberを取得
            second_Num = ws.Cells(4, colNum)
            third_Num = ws.Cells(rowNum, 9)
            
            
            randomString = ""
            For i = 1 To 2 ' 2文字のランダム文字列を生成
                randomString = randomString & Mid(charSet, Int((Len(charSet) * Rnd) + 1), 1)
            Next i
            '後ろに「1_2_3」を追加
            ws.Cells(rowNum, colNum).value = randomString & " 1_" & second_Num & "_" & third_Num
        Next colNum
    Next rowNum
End Sub



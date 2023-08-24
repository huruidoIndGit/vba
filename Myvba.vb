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


Function WorksheetExists(sheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

このコードは、指定された名前のワークシートが存在するかどうかを判定するカスタム関数 WorksheetExists を定義しています。

Function WorksheetExists(sheetName As String) As Boolean:

Function: これが関数の宣言を示すキーワードです。
WorksheetExists: 関数の名前を指定しています。任意の名前を選べますが、ここでは「WorksheetExists」という名前にしています。
(sheetName As String) As Boolean: この関数は1つの引数を受け取り、その引数の型と戻り値の型を指定しています。引数は sheetName という文字列型で、戻り値は真偽値（Boolean）です。
On Error Resume Next:

これはエラーハンドリングの一種で、エラーが発生した場合に処理を中断せず、次のステートメントに進むように指示しています。これにより、エラーが発生しても実行は続行されます。
WorksheetExists = Not Worksheets(sheetName) Is Nothing:

WorksheetExists: 戻り値として返す変数です。この変数に代入することで関数の戻り値が決まります。
Worksheets(sheetName): 引数として渡された sheetName で指定されたワークシートを取得しようとしています。
Is Nothing: オブジェクトが存在しないことを示す条件です。ここでは取得したワークシートが存在しない場合に True を返すようになっています。
Not: 論理否定演算子で、条件を反転させます。つまり、ワークシートが存在する場合は False を、存在しない場合は True を返します。
On Error GoTo 0:

エラーハンドリングを元に戻すためのステートメントです。以降のコードではエラーが発生したら中断するようになります。

Sub Main()
    Dim ws As Worksheet
    If WorksheetExists("不明なSheet名") Then
        Set ws = Worksheets("不明なSheet名")
        ' シートが存在する場合の処理
        ' ...
    Else
        ' シートが存在しない場合の処理
        ' ...
    End If
End Sub

このコードは、上記で定義した WorksheetExists 関数を使用して、ワークシートの存在を確認し、それに基づいて処理を行うサブルーチン Main を定義しています。

Sub Main():

サブルーチン（サブプロシージャ）の宣言です。コードの実行を開始するための入り口です。
Dim ws As Worksheet:

ws という変数を Worksheet 型として宣言しています。これは後でワークシートオブジェクトを格納するための変数です。
If WorksheetExists("不明なSheet名") Then:

上で定義した WorksheetExists 関数を呼び出して、指定した名前のワークシートが存在するか確認しています。もし存在する場合は条件が True になります。
Set ws = Worksheets("不明なSheet名"):

ws 変数に、指定した名前のワークシートオブジェクトを代入しています。存在しない場合でも、ws には Nothing が代入される可能性があります。
' シートが存在する場合の処理 および ' シートが存在しない場合の処理:

If 文の条件に基づいて、ワークシートが存在するかどうかに応じてそれぞれの処理を行います。必要に応じてコードを追加してください。

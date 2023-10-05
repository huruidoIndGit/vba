Option Explicit
Sub UpdateFiles()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
        
    Dim ws As Worksheet 'VBAを実行するSheet.
    Set ws = ThisWorkbook.ActiveSheet()
    ' 更新するファイルが格納されているフォルダのパスを指定します
    folderPath = SelectFiles()   '= "C:\path\to\your\folder\"
    Debug.Print folderPath
    
    folderPath = folderPath & "\"
    ' フォルダ内の最初のExcelファイルの名前を取得します
    fileName = Dir(folderPath & "*.xls*")
    
    ' フォルダ内のすべてのExcelファイルに対してループ処理を行います
    Do While fileName <> ""
        Debug.Print "File Name:" & fileName
        ' ファイルを開きます
        Set wb = Workbooks.Open(folderPath & fileName)
        ' ここで任意の操作
        Call OneFileAction(ws, wb.ActiveSheet)
        ' ファイルを保存して閉じます
        wb.Close SaveChanges:=True
        
        ' 次のファイルの名前を取得します
        fileName = Dir
    Loop
End Sub

'フォルダ選択ダイアログを表示して,そのPathを返すメソッド
Function SelectFiles() As String
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "フォルダを選択してください"
        .InitialFileName = "C:\Users\rll2p\Desktop" ' 初期表示するフォルダパスを指定します
        If .Show = -1 Then ' ユーザーがOKをクリックした場合
            folderPath = .SelectedItems(1)
            ' ここでfolderPathを使用して各ファイルに対する操作を行います
        End If
    End With
    
    ' フォルダ名を返します
    SelectFiles = folderPath
End Function

'1つのエクセルへのコピペ
Sub OneFileAction(sorceSt As Worksheet, targetSt As Worksheet)
     Dim sourceColumn As Range, targetColumn As Range

    ' コピーする列を指定します
    Set sourceColumn = sorceSt.Range("J:N")
    Set targetColumn = targetSt.Columns("K")
    sourceColumn.Copy Destination:=targetColumn
End Sub
'文字列から最初に見つかった数値（整数または小数）を抽出して返し
Function ExtractNumber(s As Variant) As Double
    Dim regEx As Object
    Dim matches As Object

    ' 正規表現オブジェクトを作成します
    Set regEx = CreateObject("VBScript.RegExp")

    ' 浮動小数点数に一致するパターンを設定します
    regEx.Pattern = "[-+]?[0-9]*\.?[0-9]+"

    ' 検索を実行します
    Set matches = regEx.Execute(s)

    ' マッチが見つかった場合、それをDoubleに変換して返します
    If matches.Count > 0 Then
        ExtractNumber = CDbl(matches.Item(0))
    Else
        ExtractNumber = 0
    End If
End Function
'指定範囲を走査する
Sub LoopThroughCells(ws As Worksheet)
    Dim i As Integer
    Dim cellValue As Variant

    ' D列の3行目から10行目までを順に走査します
    For i = 3 To 10
        Dim result As Double
        ' セルの値を取得します
        cellValue = ws.Cells(i, "D").Value
      
          If Not IsEmpty(cellValue) Then
            result = ExtractNumber(cellValue)
              Debug.Print result
          End If
    Next i
End Sub

Sub test()
    Call LoopThroughCells(ThisWorkbook.ActiveSheet)
End Sub


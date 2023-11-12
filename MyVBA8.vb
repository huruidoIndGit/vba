Sub AccessExcelBooksInHogeFolders()
    Dim rootFolder As String
    rootFolder = "C:\Users\rll2p\Desktop\ExcelTest\Root"  ' ここにrootフォルダのパスを指定
    Call ProcessFolder(rootFolder)
End Sub

Sub ProcessFolder(ByVal folderPath As String)
    Dim subFolder As Object
    Dim file As Object
    Dim folder As Object
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    For Each subFolder In folder.SubFolders
        If subFolder.Name = "hoge" Then
            For Each file In subFolder.Files
                If fso.GetExtensionName(file) = "csv" Then
                    'Workbooks.Open file.Path
                    Debug.Print file.Path
                End If
            Next file
        Else
            Call ProcessFolder(subFolder.Path)
        End If
    Next subFolder
End Sub



Sub FindAndAccess()
    Dim rng As range
    Dim cell As range
    Dim ws As Worksheet
    Set ws = wb.ActiveSheet

    ' "hoge"を含むセルを見つける
    Set rng = ws.range("A:A").Find("hoge", LookIn:=xlValues)

    If Not rng Is Nothing Then
        ' "hoge"の次のセルから空白行まで繰り返しアクセスする
        Set cell = rng.Offset(1, 0)
        Do While cell.Value <> ""
            ' ここに処理を書く
            Debug.Print cell.Address & ": " & cell.Value ' 例：セルのアドレスと値を出力
            Set cell = cell.Offset(1, 0)
        Loop
    Else
        MsgBox "hogeが見つかりませんでした。"
    End If
End Sub




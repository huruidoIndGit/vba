Sub GetFileNamesInFolder()
    ' 指定するフォルダのパスを定数として指定します
    Const FolderPath As String = "C:\Users\rll2p\Desktop\test\"
    
    Dim FileSystem As Object
    Dim Folder As Object
    Dim File As Object
    Dim FileName As String
    
    ' スクリプトのファイル システム オブジェクトを作成します
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    ' 指定したフォルダへの参照を取得します
    Set Folder = FileSystem.GetFolder(FolderPath)
    
    ' フォルダ内の各ファイルに対して処理を行います
    For Each File In Folder.Files
        ' ファイル名を取得します
        FileName = File.Name
        ' 取得したファイル名を出力します（ここでは即座にメッセージボックスで表示）
        
        
        MsgBox FileName
    Next File
    
    ' オブジェクトを解放します
    Set File = Nothing
    Set Folder = Nothing
    Set FileSystem = Nothing
End Sub


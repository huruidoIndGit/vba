Sub SaveAsUTF8WithoutBOM() 'filePath As String, newFilePath As String)
    Const filePath As String = "C:\Users\rll2p\Desktop\Book2.csv"
    Const newFilePath As String = "C:\Users\rll2p\Desktop\ExcelTest\Hoge"


    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim file As Object
    Set file = fso.OpenTextFile(filePath, 1)
    
    Dim content As String
    content = file.ReadAll
    
    file.Close
    
    content = Replace(content, """", "")
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Charset = "utf-8"
    stream.WriteText content
    stream.Position = 0
    stream.Type = 1
    stream.Position = 3
    
    Dim newStream As Object
    Set newStream = CreateObject("ADODB.Stream")
    newStream.Type = 1
    newStream.Open
    stream.CopyTo newStream
    newStream.SaveToFile newFilePath & ".txt", 2 ' 2 = overwrite
    newStream.Close
    stream.Close
End Sub


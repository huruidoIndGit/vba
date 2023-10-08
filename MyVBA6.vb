
Function IsInRange(range As Variant, number As Double) As Boolean
    IsInRange = number >= range(0) And number <= range(1)
End Function

'Sub CheckNumberInRange()
'    Dim i As Integer
'
'    ' 正規表現オブジェクトを作成
'    Dim regEx As New RegExp
'    regEx.Pattern = "([-]?[0-9.±~]+)([a-zA-Z]*).*"
'    Debug.Print WorksheetFunction.CountA(Columns("D"))
'    For i = 3 To WorksheetFunction.CountA(Columns("D")) + 4
'        If Not IsEmpty(range("D" & i)) And regEx.test(range("D" & i).Value) Then
''            range("G" & i).Value = IsInRange(ExtractRange(range("D" & i)), range("F" & i).Value)
'            Dim rng As Variant
'            rng = ExtractRange(range("D" & i))
'            range("E" & i).Value = IsInRange(rng, range("F" & i).Value)
'            ' K列に条件式を追加
'            range("K" & i).Formula = "=IF(F" & i & "="""","""",(IF(AND(F" & i & ">=" & rng(0) & ",F" & i & "<=" & rng(1) & "),1,0)))"
'        End If
'    Next i
'End Sub
Sub CheckNumberInRange()
    Dim i As Integer
    Dim j As Integer
    Dim col As String
    ' 正規表現オブジェクトを作成
    Dim regEx As New RegExp
    regEx.Pattern = "([-]?[0-9.±~]+)([a-zA-Z]*).*"
   ' Debug.Print WorksheetFunction.CountA(Columns("D"))
    For i = 3 To WorksheetFunction.CountA(Columns("D")) + 3
        If Not IsEmpty(range("D" & i)) And regEx.test(range("D" & i).Value) Then
            Dim rng As Variant
            rng = ExtractRange(range("D" & i))
            For j = 6 To 9 ' F列からI列まで
                col = Chr(64 + j)
               ' range(col & i).Value = IsInRange(rng, range(col & i).Value)
                ' K列からN列に条件式を追加
                range(Chr(64 + j + 5) & i).Formula = "=IF(" & col & i & "="""","""",(IF(AND(" & col & i & ">=" & rng(0) & "," & col & i & "<=" & rng(1) & "),1,0)))" 'Trueなら1,Falseなら0
            Next j
        End If
    Next i
End Sub


'参照設定ダイアログで「Microsoft VBScript Regular Expressions 5.5」にチェック付ける必要がある
Function ExtractRange(cell As range) As Variant
    Debug.Print cell.Address
    Dim str As String
    str = cell.Value
    Dim unit As String
    Dim number As String
    Dim splitStr() As String
    
    ' 正規表現オブジェクトを作成
    Dim regEx As New RegExp
    regEx.Pattern = "([-]?[0-9.±~]+)([a-zA-Z]*).*"
    
    ' 正規表現でマッチング
    Dim matches As MatchCollection
    Set matches = regEx.Execute(str)
    
    If matches.Count > 0 Then
        ' 数値部分と単位部分を取得
        number = matches.Item(0).SubMatches(0)
        unit = matches.Item(0).SubMatches(1)
        Debug.Print "unit:" & unit & "    number:" & number
        If InStr(number, "±") > 0 Then
            splitStr = Split(number, "±")
            ExtractRange = Array(CDbl(splitStr(0)) - CDbl(splitStr(1)), CDbl(splitStr(0)) + CDbl(splitStr(1)))
        ElseIf InStr(number, "~") > 0 Then
            splitStr = Split(number, "~")
            ExtractRange = Array(CDbl(splitStr(0)), CDbl(splitStr(1)))
        ElseIf InStr(str, "以下") > 0 Then
            ExtractRange = Array(-1E+307, CDbl(Replace(number, "以下", "")))
        ElseIf InStr(str, "以上") > 0 Then
            ExtractRange = Array(CDbl(Replace(number, "以上", "")), 1E+307)
        End If
    End If
End Function

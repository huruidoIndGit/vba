Option Explicit

'指定したSheetから図形をコピペする&グループ化&命名
Public Sub Run()
    Dim srcWb As Workbook
    Dim destWb As Workbook
    Dim srcWs As Worksheet
    Dim destWs As Worksheet
    Dim sheetIndex As Integer
    Dim visibleSheetIndex As Integer
    Dim pasteStartCell As Range
    Const targetRange As String = "C1:G10"      ' 図形が収まってる範囲
    Const pasteStartAddress As String = "C5"    ' 貼り付け開始セル
    Const pasteInterval As Integer = 5          ' 貼り付け間隔（列数）
    Const maxShapesPerRow As Integer = 3        ' 1行に貼り付ける最大図形数
    Const rowInterval As Integer = 10            ' 行方向の貼り付け間隔（行数）

    ' コピー元のブックを取得
    Set srcWb = 開いているブックの取得Module.GetOtherOpenWorkbook 'もう一つのBookを取得
    Debug.Print srcWb.Name

    ' 貼り付け先のブックを取得
    Set destWb = ThisWorkbook
    Set pasteStartCell = destWb.Sheets(1).Range(pasteStartAddress)
    
    ' 各シートから図形を取得してグループ化し、貼り付け
    visibleSheetIndex = 1
    For sheetIndex = 1 To srcWb.Sheets.Count
        Set srcWs = srcWb.Sheets(sheetIndex)
        If srcWs.Visible = xlSheetVisible Then
            Set destWs = destWb.Sheets(1) ' 貼り付け先のシートを指定
            GroupAndCopyShapes srcWs, destWs, targetRange, pasteStartCell, visibleSheetIndex, pasteInterval, maxShapesPerRow, rowInterval
            visibleSheetIndex = visibleSheetIndex + 1
        End If
    Next sheetIndex
End Sub

' 図形をグループ化してコピーするメソッド
Private Sub GroupAndCopyShapes(srcWs As Worksheet, destWs As Worksheet, targetRange As String, pasteStartCell As Range, sheetIndex As Integer, pasteInterval As Integer, maxShapesPerRow As Integer, rowInterval As Integer)
    Dim shp As Shape
    Dim shpRange As Range
    Dim shpGroup As Shape
    Dim shpArray() As String
    Dim i As Integer
    Dim msg As String
    Dim pasteCell As Range
    Dim rowIndex As Integer
    Dim colIndex As Integer
    
    ' 図形を検索する範囲を設定
    Set shpRange = srcWs.Range(targetRange)
    
    ' 図形を配列に追加
    i = 0
    msg = vbCrLf
    For Each shp In srcWs.Shapes
        If Not Intersect(shpRange, shp.TopLeftCell) Is Nothing Then
            ReDim Preserve shpArray(i)
            shpArray(i) = shp.Name
            msg = msg & shp.Name & " (" & shp.Type & ") " & vbCrLf ' 図形の名前と種類をメッセージに追加
            i = i + 1
        End If
    Next shp
    
    ' 図形の数をメッセージに追加
    msg = msg & "図形の数==>>: " & i
    
    ' 図形の名前と数を表示（確認用）
    Debug.Print msg
    
    ' 図形をグループ化
    If i > 1 Then
        Set shpGroup = srcWs.Shapes.Range(shpArray).Group
        shpGroup.Name = "@" & sheetIndex
        
        ' 貼り付け先のセルを計算
        rowIndex = (sheetIndex - 1) \ maxShapesPerRow
        colIndex = (sheetIndex - 1) Mod maxShapesPerRow
        Set pasteCell = pasteStartCell.Offset(rowIndex * rowInterval, colIndex * pasteInterval)
        
        ' グループ化した図形をコピー
        shpGroup.Copy
        destWs.Paste Destination:=pasteCell
        
        ' 貼り付けた図形の名前を変更
        destWs.Shapes(destWs.Shapes.Count).Name = "@" & sheetIndex
        
        ' 一時的な図形を削除
        shpGroup.Ungroup
    ElseIf i = 1 Then
        ' 図形が1つだけの場合はグループ化せずにコピー
        srcWs.Shapes(shpArray(0)).Copy
        rowIndex = (sheetIndex - 1) \ maxShapesPerRow
        colIndex = (sheetIndex - 1) Mod maxShapesPerRow
        Set pasteCell = pasteStartCell.Offset(rowIndex * rowInterval, colIndex * pasteInterval)
        destWs.Paste Destination:=pasteCell
        
        ' 貼り付けた図形の名前を変更
        destWs.Shapes(destWs.Shapes.Count).Name = "@" & sheetIndex
    Else
        Debug.Print "指定した範囲内に図形が見つかりませんでした: " & srcWs.Name
    End If
End Sub


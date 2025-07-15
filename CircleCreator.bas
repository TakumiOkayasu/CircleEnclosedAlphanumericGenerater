Attribute VB_Name = "CircleCreator"
Option Explicit

' 定数定義
Private Const CIRCLE_HEIGHT_CM As Single = 0.82!
Private Const CIRCLE_LINE_WEIGHT As Single = 1.5!
Private Const POINTS_PER_INCH As Single = 72!
Private Const CM_PER_INCH As Single = 2.54!

' センチメートルをポイントに変換
Private Function CmToPoints(ByVal cm As Single) As Single
    CmToPoints = (cm / CM_PER_INCH) * POINTS_PER_INCH
End Function

' 数字入りの正円を作成する基本的なプロシージャ
Sub CreateNumberedCircle(number As Long, x As Single, y As Single)
    Dim ws As Worksheet
    Dim shpCircle As Shape
    Dim leftPos As Single
    Dim topPos As Single
    Dim circleSize As Single
    Dim numberText As String
    
    ' アクティブシートを設定
    Set ws = ActiveSheet
    
    ' 位置と数字を設定
    leftPos = x! ' 左位置（ポイント）
    topPos = y! ' 上位置（ポイント）
    
    ' 0.82cmをポイントに変換
    circleSize = CmToPoints(CIRCLE_HEIGHT_CM)
    
    ' 正円（Oval）を作成
    Set shpCircle = ws.Shapes.AddShape(msoShapeOval, _
        leftPos, _
        topPos, _
        circleSize, _
        circleSize _
    )
    
    ' 円のスタイルを設定
    With shpCircle
        ' 塗りつぶしなし（透明）
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        
        ' 線の設定
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Line.Weight = CIRCLE_LINE_WEIGHT ' 1.5ポイント
        
        ' 縦横比を固定
        .LockAspectRatio = msoTrue
        
        ' 図形の名前を設定
        .Name = "NumberCircle_" & number

        ' テキストを追加・垂直配置を中央に
        With .TextFrame2
            .WordWrap = msoFalse ' 文字を折り返さない
            .TextRange.Characters.Text = CStr(number)
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 0!
            .MarginRight = 0!
            .MarginTop = 0!
            .MarginBottom = 0!
        End With
        
        ' テキストの書式設定
        With .TextFrame2.TextRange
            .Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
            .Font.size = 16
            .Font.Name = "Calibri 本文"
            .Font.Bold = msoFalse
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    
End Sub

Sub CreateMultipleNumberedCircle()

    Dim i As Long
    Dim startX As Long
    Dim startY As Long
    Dim offsetPixels As Long
    Dim shapeNames() As String
    Dim ws As Worksheet
    Dim beginCount As Long
    Dim endCount As Long
    
    ' アクティブシートを設定
    Set ws = ActiveSheet
    
    ' 数字が入力してあれば
    If Application.WorksheetFunction.IsNumber(ws.Range("G3").value) Then
        beginCount = ws.Range("G3").value
    Else
        MsgBox ("半角数字を入力してね")
        Exit Sub
    End If
    
    If Application.WorksheetFunction.IsNumber(ws.Range("G4").value) Then
        endCount = ws.Range("G4").value
    Else
        MsgBox ("半角数字を入力してね")
        Exit Sub
    End If
    
    ' 開始位置
    startX = 100!
    startY = 100!
    
    ' オフセット（30ピクセル）
    offsetPixels = 10!
    
    ' 図形名を格納する配列を準備
    ReDim shapeNames(1 To endCount)
    
    ' 1からoutputCountまでの円を作成（30ピクセルずつ右下に配置）
    For i = beginCount To endCount
        CreateNumberedCircle i, _
            startX + (i - 1) * offsetPixels, _
            startY + (i - 1) * offsetPixels
        ' 作成した図形の名前を配列に格納
        shapeNames(i) = "NumberCircle_" & i
    Next
    
    ' すべての生成した図形を選択
    ws.Shapes.Range(shapeNames).Select
End Sub

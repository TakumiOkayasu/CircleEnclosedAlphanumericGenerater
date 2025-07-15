Attribute VB_Name = "CircleCreator"
Option Explicit

' �萔��`
Private Const CIRCLE_HEIGHT_CM As Single = 0.82!
Private Const CIRCLE_LINE_WEIGHT As Single = 1.5!
Private Const POINTS_PER_INCH As Single = 72!
Private Const CM_PER_INCH As Single = 2.54!

' �Z���`���[�g�����|�C���g�ɕϊ�
Private Function CmToPoints(ByVal cm As Single) As Single
    CmToPoints = (cm / CM_PER_INCH) * POINTS_PER_INCH
End Function

' ��������̐��~���쐬�����{�I�ȃv���V�[�W��
Sub CreateNumberedCircle(number As Long, x As Single, y As Single)
    Dim ws As Worksheet
    Dim shpCircle As Shape
    Dim leftPos As Single
    Dim topPos As Single
    Dim circleSize As Single
    Dim numberText As String
    
    ' �A�N�e�B�u�V�[�g��ݒ�
    Set ws = ActiveSheet
    
    ' �ʒu�Ɛ�����ݒ�
    leftPos = x! ' ���ʒu�i�|�C���g�j
    topPos = y! ' ��ʒu�i�|�C���g�j
    
    ' 0.82cm���|�C���g�ɕϊ�
    circleSize = CmToPoints(CIRCLE_HEIGHT_CM)
    
    ' ���~�iOval�j���쐬
    Set shpCircle = ws.Shapes.AddShape(msoShapeOval, _
        leftPos, _
        topPos, _
        circleSize, _
        circleSize _
    )
    
    ' �~�̃X�^�C����ݒ�
    With shpCircle
        ' �h��Ԃ��Ȃ��i�����j
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        
        ' ���̐ݒ�
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Line.Weight = CIRCLE_LINE_WEIGHT ' 1.5�|�C���g
        
        ' �c������Œ�
        .LockAspectRatio = msoTrue
        
        ' �}�`�̖��O��ݒ�
        .Name = "NumberCircle_" & number

        ' �e�L�X�g��ǉ��E�����z�u�𒆉���
        With .TextFrame2
            .WordWrap = msoFalse ' ������܂�Ԃ��Ȃ�
            .TextRange.Characters.Text = CStr(number)
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 0!
            .MarginRight = 0!
            .MarginTop = 0!
            .MarginBottom = 0!
        End With
        
        ' �e�L�X�g�̏����ݒ�
        With .TextFrame2.TextRange
            .Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
            .Font.size = 16
            .Font.Name = "Calibri �{��"
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
    
    ' �A�N�e�B�u�V�[�g��ݒ�
    Set ws = ActiveSheet
    
    ' ���������͂��Ă����
    If Application.WorksheetFunction.IsNumber(ws.Range("G3").value) Then
        beginCount = ws.Range("G3").value
    Else
        MsgBox ("���p��������͂��Ă�")
        Exit Sub
    End If
    
    If Application.WorksheetFunction.IsNumber(ws.Range("G4").value) Then
        endCount = ws.Range("G4").value
    Else
        MsgBox ("���p��������͂��Ă�")
        Exit Sub
    End If
    
    ' �J�n�ʒu
    startX = 100!
    startY = 100!
    
    ' �I�t�Z�b�g�i30�s�N�Z���j
    offsetPixels = 10!
    
    ' �}�`�����i�[����z�������
    ReDim shapeNames(1 To endCount)
    
    ' 1����outputCount�܂ł̉~���쐬�i30�s�N�Z�����E���ɔz�u�j
    For i = beginCount To endCount
        CreateNumberedCircle i, _
            startX + (i - 1) * offsetPixels, _
            startY + (i - 1) * offsetPixels
        ' �쐬�����}�`�̖��O��z��Ɋi�[
        shapeNames(i) = "NumberCircle_" & i
    Next
    
    ' ���ׂĂ̐��������}�`��I��
    ws.Shapes.Range(shapeNames).Select
End Sub

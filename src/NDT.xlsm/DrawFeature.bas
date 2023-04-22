Attribute VB_Name = "DrawFeature"
Option Explicit
Const SIZE = 60
Const Y_DISTANCE = 20
Const X_DISTANCE = 10
Const X_OFFSET = 10
Const Y_OFFSET = 125
Const CONNECTOR_COLOR = XlRgbColor.rgbDimGray

Private Enum Directions
    North = 1
    NorthWest
    West
    SouthWest
    South
    SouthEast
    East
    NorthEast
End Enum

Public Sub EntryPoint()
    If ConfigSheet.LockMacro Then
        MsgBox "このマクロは既存のデザインに影響を及ぼす可能性があるため、現在ロックされています。" & vbNewLine & "リスクを承知のうえでロックを解除するにはConfigシートのC4セルをFalseに書き換えてください。", vbExclamation
        Exit Sub
    End If
    Application.Run Application.Caller
End Sub

Public Sub EntryPointOperational()
    Application.Run Application.Caller
End Sub

Private Sub Btn_PlotTasks()
    Application.ScreenUpdating = False
    Call RemoveUnregisteredOvals
    Dim x As Double: x = X_DISTANCE
    Dim y As Double: y = Y_DISTANCE
    Dim n As Node
    x = 0

    With DrawSheet.Rows(10)
        .ClearOutline
        .Group
    End With
    For Each n In ScheduleSheet.GetTaskListAsNodes
        Dim sh As Shape
        Set sh = n.FindShape
        If sh Is Nothing Then
            n.ShapeObjectName = CreateNodeShape(SIZE, x, y, n.TaskTitle).Name
            n.TaskListRange.Offset(0, 5).Value = n.ShapeObjectName
            x = x + SIZE + X_DISTANCE
        Else
            sh.TextFrame2.TextRange.Text = OptimizeTextReturn(n.TaskTitle, 5)
        End If
    Next
    DrawSheet.Select
    Application.ScreenUpdating = True
End Sub

Private Function CreateNodeShape(SIZE, pos_x, pos_y, task_title As String) As Shape
    Dim s As Shape
    Set s = DrawSheet.Shapes.AddShape(msoShapeOval, pos_x + X_OFFSET, pos_y + Y_OFFSET, SIZE, SIZE)
    s.Fill.ForeColor.RGB = XlRgbColor.rgbLavender
    s.Line.Visible = msoFalse
    s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbBlack
    With s.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .HorizontalAnchor = msoAnchorNone
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    s.TextFrame2.TextRange.Text = OptimizeTextReturn(task_title, 5)
    s.TextFrame2.WordWrap = msoFalse
    s.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
    s.TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow
    Set CreateNodeShape = s
End Function

Private Function OptimizeTextReturn(original_text, normal_width) As String
    Dim w As Integer: w = normal_width
    Dim h As Integer: h = Round(Len(original_text) / normal_width + 0.4, 0)
    
    Do Until h <= w
        w = w + 1
        h = Round(Len(original_text) / w + 0.4, 0)
    Loop

    Dim result_string As String: result_string = ""
    Dim rest_string As String: rest_string = original_text
    
    Do While Len(rest_string) > 0
        result_string = result_string & Left(rest_string, w) & vbCrLf
        rest_string = Mid(rest_string, w + 1)
    Loop
    
    'I'm still investigation that why minus 2 works. I just  wanted to remove the last vbCrLf. IS vbCrLf two letter?"
    If Len(result_string) >= 2 Then
        result_string = Left(result_string, Len(result_string) - 2)
    End If
    
    OptimizeTextReturn = result_string
End Function

Private Sub Btn_RemoveAllShapse()
    Dim sh As Shape
    For Each sh In DrawSheet.Shapes
        If sh.Type <> msoFormControl Then
            sh.Delete
        End If
    Next
End Sub

Private Sub RemoveUnregisteredOvals()
    Dim ov As Oval
    Dim n As Node
    Dim nn As Nodes: Set nn = ScheduleSheet.GetTaskListAsNodes
    For Each ov In DrawSheet.Ovals
        For Each n In nn
            If ov.Name = n.ShapeObjectName Then GoTo Continue
        Next
        ov.Delete
Continue:
    Next
    RemoveDisconnection
End Sub

Private Sub RemoveDisconnection()
    Dim sh As Shape
    For Each sh In DrawSheet.Shapes
        ' This magic number -2 is just taken from an inspection result.
        ' It's not yet logically confirmed that -2 always indicate the connector in this usage.
        If sh.Type = msoAutoShape And sh.AutoShapeType = -2 Then
            If sh.ConnectorFormat.BeginConnected And sh.ConnectorFormat.EndConnected Then
                'Do Nothing
            Else
                sh.Delete
            End If
        End If
    Next
End Sub

Private Sub Btn_FindDisconnection()
    Dim sh As Shape
    For Each sh In DrawSheet.Shapes
        ' This magic number -2 is just taken from an inspection result.
        ' It's not yet logically confirmed that -2 always indicate the connector in this usage.
        If sh.Type = msoAutoShape And sh.AutoShapeType = -2 Then
            If sh.ConnectorFormat.BeginConnected And sh.ConnectorFormat.EndConnected Then
                sh.Line.ForeColor.RGB = CONNECTOR_COLOR
            Else
                sh.Line.ForeColor.RGB = vbRed
            End If
        End If
    Next
End Sub

Private Sub Btn_RemoveConnections()
    Dim sh As Shape
    For Each sh In DrawSheet.Shapes
        ' This magic number -2 is just taken from an inspection result.
        ' It's not yet logically confirmed that -2 always indicate the connector in this usage.
        If sh.Type = msoAutoShape And sh.AutoShapeType = -2 Then
            sh.Delete
        End If
    Next
End Sub

Private Sub Btn_SwapNodeLocation()
    If getOvalCollection_SE.Count <> 2 Then
        MsgBox "Kindly select just 2 nodes.", vbInformation
        Exit Sub
    End If
    
    Dim sh1 As Shape
    Dim sh2 As Shape
    Set sh1 = Selection.ShapeRange(1)
    Set sh2 = Selection.ShapeRange(2)
    
    Dim tt As Single
    tt = sh1.Top
    Dim ll As Single
    ll = sh1.Left
    sh1.Top = sh2.Top
    sh1.Left = sh2.Left
    sh2.Top = tt
    sh2.Left = ll
End Sub

Private Sub Btn_OrderNodeVertical()
    'To keep selection order, store shapes to a Collection.
    Dim c As Collection
    Set c = getOvalCollection_SE
    If c.Count < 2 Then
        MsgBox "Kindly select 2 nodes at least.", vbInformation
        Exit Sub
    End If
    
    Dim sh As Shape
    Selection.ShapeRange.Group.Select
    Dim leftEdge  As Single
    leftEdge = Round(Selection.ShapeRange.Left, 3)
    Dim topEdge As Single
    topEdge = Selection.ShapeRange.Top
    Selection.ShapeRange.Ungroup
    
    Dim n As Integer
    For Each sh In c
        sh.Left = leftEdge
        sh.Top = topEdge + (SIZE + Y_DISTANCE) * n
        n = n + 1
    Next
End Sub

Private Sub Btn_ConnectStreight()
    With getOvalCollection_SE
        Dim i As Integer
        For i = 1 To .Count - 1
            Call ConnectArrow( _
                src_oval:=.Item(i), _
                dst_oval:=.Item(i + 1))
        Next
    End With
End Sub

Private Sub Btn_ConnectSplit()
    With getOvalCollection_SE
    Dim i As Integer
        For i = 2 To .Count
            Call ConnectArrow( _
                src_oval:=.Item(1), _
                dst_oval:=.Item(i))
        Next
    End With
End Sub

Private Sub Btn_ConnectMarge()
    With getOvalCollection_SE
        Dim i As Integer
        For i = 1 To .Count - 1
            Call ConnectArrow( _
                src_oval:=.Item(i), _
                dst_oval:=.Item(.Count))
        Next
    End With
End Sub

'Postfix _SE means that it has Side Effects
Private Function getOvalCollection_SE() As Collection
    Set getOvalCollection_SE = New Collection
    If TypeName(Selection) = "Range" Then Exit Function
    
    Dim shp As Shape
    For Each shp In Selection.ShapeRange
        If shp.AutoShapeType = msoShapeOval Then
            getOvalCollection_SE.Add shp
        End If
    Next
End Function

Private Sub ConnectArrow(src_oval As Shape, dst_oval As Shape)
        Dim arrow As Shape
        Set arrow = DrawSheet.Shapes.AddConnector(msoConnectorStraight, 0, 0, 100, 100)
        
        arrow.Line.ForeColor.RGB = CONNECTOR_COLOR
        arrow.Line.EndArrowheadStyle = msoArrowheadTriangle
        
        arrow.ConnectorFormat.BeginConnect _
            ConnectedShape:=src_oval, _
            ConnectionSite:=Directions.East
        
        arrow.ConnectorFormat.EndConnect _
            ConnectedShape:=dst_oval, _
            ConnectionSite:=Directions.West
End Sub

Private Sub Btn_NumberingNodes()
    Dim i As Integer: i = 0
    Dim sh As Shape
    Dim nn As Nodes: Set nn = ScheduleSheet.GetTaskListAsNodes
    Dim n As Node
    For Each sh In Selection.ShapeRange
        If sh.Type = msoAutoShape And sh.AutoShapeType = 9 Then
            Dim tmpTaskTitle As String: tmpTaskTitle = Replace(sh.TextFrame2.TextRange.Text, vbLf, "")
            tmpTaskTitle = i & "." & RemoveNumberPrefix(tmpTaskTitle)
            sh.TextFrame2.TextRange.Text = OptimizeTextReturn(tmpTaskTitle, 5)
            For Each n In nn
                If sh.Name = n.ShapeObjectName Then
                    n.TaskListRange.Offset(0, -1).Value = CLng(i)
                End If
            Next
            i = i + 1
        End If
    Next
End Sub

Private Sub Btn_DeNumberingAllNodes()
    Dim sh As Shape
    For Each sh In DrawSheet.Shapes
        If sh.Type = msoAutoShape And sh.AutoShapeType = 9 Then
            Dim tmpTaskTitle As String: tmpTaskTitle = Replace(sh.TextFrame2.TextRange.Text, vbLf, "")
            tmpTaskTitle = RemoveNumberPrefix(tmpTaskTitle)
            sh.TextFrame2.TextRange.Text = OptimizeTextReturn(tmpTaskTitle, 5)
        End If
    Next

    Dim n As Node
    For Each n In ScheduleSheet.GetTaskListAsNodes
        n.TaskListRange.Offset(0, -1).Value = ""
    Next
End Sub

Private Function RemoveNumberPrefix(tmp_str As String) As String
    If IsNumeric(Split(tmp_str, ".")(0)) Then
        RemoveNumberPrefix = Mid(tmp_str, InStr(1, tmp_str, ".") + 1)
    Else
        RemoveNumberPrefix = tmp_str
    End If
End Function

Private Sub Btn_SetCompletedIcon()
    On Error Resume Next
    IconSheet.ChartObjects("CompletedIcon").Chart.Export Environ("temp") & "\NDT_CompletedIcon.bmp"
    Selection.ShapeRange.Fill.UserPicture Environ("temp") & "\NDT_CompletedIcon.bmp"
    On Error GoTo 0
End Sub

Private Sub Btn_SetCancelledIcon()
    On Error Resume Next
    IconSheet.ChartObjects("CancelledIcon").Chart.Export Environ("temp") & "\NDT_CancelledIcon.bmp"
    Selection.ShapeRange.Fill.UserPicture Environ("temp") & "\NDT_CancelledIcon.bmp"
    On Error GoTo 0
End Sub

Private Sub Btn_SetInProgressIcon()
    Dim percentage: percentage = InputBox("進捗率(%)を1～99の整数で入力してください。")
    
    If Not IsNumeric(percentage) Then
        MsgBox "数値以外のものが入力されました。やり直してください。", vbExclamation, "エラー"
        Exit Sub
    End If

    If CDbl(percentage) <> CInt(percentage) Then
        MsgBox "少数は入力できません。やり直してください。", vbExclamation, "エラー"
        Exit Sub
    End If
    
    If percentage < 1 Or percentage > 99 Then
        MsgBox "範囲外の数値が入力されました。やり直してください。", vbExclamation, "エラー"
        Exit Sub
    End If
    
    With IconSheet.ChartObjects("InProgressIcon").Chart.Shapes(1)
        .Adjustments.Item(1) = -90
        .Adjustments.Item(2) = ((360 / 100) * CInt(percentage)) - 90
    End With
    On Error Resume Next
    IconSheet.ChartObjects("InProgressIcon").Chart.Export Environ("temp") & "\NDT_InProgressIcon.bmp"
    Selection.ShapeRange.Fill.UserPicture Environ("temp") & "\NDT_InProgressIcon.bmp"
    On Error GoTo 0
End Sub


Private Sub Btn_ClearIcon()
    On Error Resume Next
    With Selection.ShapeRange.Fill
        .Solid
        .ForeColor.RGB = rgbLavender
    End With
    On Error GoTo 0
End Sub

Private Sub Btn_LockMacros()
    ConfigSheet.LockMacro = True
End Sub

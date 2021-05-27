Attribute VB_Name = "DrawFeature"
Option Explicit
Const SIZE = 60
Const Y_DISTANCE = 20
Const X_DISTANCE = 10
Const X_OFFSET = 50
Const Y_OFFSET = 150
Const CONNECTOR_COLOR = XlRgbColor.rgbDimGray

Sub DrawTaskAsNode()
    Dim x As Double: x = X_DISTANCE
    Dim y As Double: y = Y_DISTANCE
    Dim n As Node
    'Call CreateNodeShape(SIZE, x, y, "START")
    x = x + SIZE + X_DISTANCE
    For Each n In TaskListSheet.GetTaskListAsNodes
        Dim sh As Shape
        Set sh = n.FindShape
        If sh Is Nothing Then
            n.ShapeObjectName = CreateNodeShape(SIZE, x, y, n.TaskTitle).Name
            n.TaskListRange.Offset(0, 1).Value = n.ShapeObjectName
            x = x + SIZE + X_DISTANCE
        Else
            sh.TextFrame2.TextRange.Text = OptimizeTextReturn(n.TaskTitle, 5)
        End If
        
    Next
    'Call CreateNodeShape(SIZE, x, y, "END")
End Sub

Function CreateNodeShape(SIZE, pos_x, pos_y, task_title As String) As Shape
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

Function OptimizeTextReturn(original_text, normal_width) As String
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
    result_string = Left(result_string, Len(result_string) - 2)
    
    OptimizeTextReturn = result_string
End Function

Sub RemoveAllShapse()
    Dim sh As Shape
    For Each sh In DrawSheet.Shapes
        If sh.Type <> msoFormControl Then
            sh.Delete
        End If
    Next
End Sub

Sub RemoveUnregisteredOvals()
    Dim ov As Oval
    Dim n As Node
    Dim nn As Nodes: Set nn = TaskListSheet.GetTaskListAsNodes
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

Sub FindDisconnection()
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

Sub RemoveConnections()
    Dim sh As Shape
    For Each sh In DrawSheet.Shapes
        ' This magic number -2 is just taken from an inspection result.
        ' It's not yet logically confirmed that -2 always indicate the connector in this usage.
        If sh.Type = msoAutoShape And sh.AutoShapeType = -2 Then
            sh.Delete
        End If
    Next
End Sub

Sub SwapNodeLocation()
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

Sub OrderNodeVertical()
    'To keep selection order, store shapes to a Collection.
    Dim c As Collection
    Set c = New Collection
    Dim sh As Shape
    For Each sh In Selection.ShapeRange
        c.Add sh
    Next

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

Sub ConnectStreight()
    Dim c As Collection: Set c = New Collection
    Dim sh As Shape
    For Each sh In Selection.ShapeRange
        c.Add sh
    Next
    
    Dim sh2 As Shape
    Dim i As Integer
    For i = 1 To c.Count - 1
        Set sh = c.Item(i)
        Set sh2 = c.Item(i + 1)
        
        Dim cn As Shape
        Set cn = DrawSheet.Shapes.AddConnector(msoConnectorStraight, 0, 0, 100, 100)
        cn.Line.ForeColor.RGB = CONNECTOR_COLOR
        cn.Line.EndArrowheadStyle = msoArrowheadTriangle
        cn.ConnectorFormat.BeginConnect sh, 7
        cn.ConnectorFormat.EndConnect sh2, 3
    Next
End Sub

Sub ConnectSplit()
    Dim c As Collection: Set c = New Collection
    Dim sh As Shape
    For Each sh In Selection.ShapeRange
        c.Add sh
    Next
    
    Dim sh2 As Shape
    Dim i As Integer
    For i = 2 To c.Count
        Set sh = c.Item(1)
        Set sh2 = c.Item(i)
        
        Dim cn As Shape
        Set cn = DrawSheet.Shapes.AddConnector(msoConnectorStraight, 0, 0, 100, 100)
        cn.Line.ForeColor.RGB = CONNECTOR_COLOR
        cn.Line.EndArrowheadStyle = msoArrowheadTriangle
        cn.ConnectorFormat.BeginConnect sh, 7
        cn.ConnectorFormat.EndConnect sh2, 3
    Next
End Sub

Sub ConnectMarge()
    Dim c As Collection: Set c = New Collection
    Dim sh As Shape
    For Each sh In Selection.ShapeRange
        c.Add sh
    Next
    
    Dim sh2 As Shape
    Dim i As Integer
    For i = 1 To c.Count - 1
        Set sh = c.Item(i)
        Set sh2 = c.Item(c.Count)
        
        Dim cn As Shape
        Set cn = DrawSheet.Shapes.AddConnector(msoConnectorStraight, 0, 0, 100, 100)
        cn.Line.ForeColor.RGB = CONNECTOR_COLOR
        cn.Line.EndArrowheadStyle = msoArrowheadTriangle
        cn.ConnectorFormat.BeginConnect sh, 7
        cn.ConnectorFormat.EndConnect sh2, 3
    Next
End Sub

Sub NumberingNodes()
    Dim i As Integer: i = 0
    Dim sh As Shape
    Dim nn As Nodes: Set nn = TaskListSheet.GetTaskListAsNodes
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

Sub DeNumberingAllNodes()
    Dim sh As Shape
    For Each sh In DrawSheet.Shapes
        If sh.Type = msoAutoShape And sh.AutoShapeType = 9 Then
            Dim tmpTaskTitle As String: tmpTaskTitle = Replace(sh.TextFrame2.TextRange.Text, vbLf, "")
            tmpTaskTitle = RemoveNumberPrefix(tmpTaskTitle)
            sh.TextFrame2.TextRange.Text = OptimizeTextReturn(tmpTaskTitle, 5)
        End If
    Next

    Dim n As Node
    For Each n In TaskListSheet.GetTaskListAsNodes
        n.TaskListRange.Offset(0, -1).Value = ""
    Next
End Sub

Function RemoveNumberPrefix(tmp_str As String) As String
    If IsNumeric(Split(tmp_str, ".")(0)) Then
        RemoveNumberPrefix = Mid(tmp_str, InStr(1, tmp_str, ".") + 1)
    Else
        RemoveNumberPrefix = tmp_str
    End If
End Function

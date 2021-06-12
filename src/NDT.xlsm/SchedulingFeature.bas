Attribute VB_Name = "SchedulingFeature"
Option Explicit
Const CONNECTOR_COLOR = XlRgbColor.rgbDimGray
Private Function CheckAllNodeNumbered() As Boolean
    Dim dic As Dictionary: Set dic = New Dictionary
    Dim ov As Oval
    Dim k As String
    For Each ov In DrawSheet.Ovals
        k = Split(ov.Text, ".")(0)
        If Not IsNumeric(k) Then
            CheckAllNodeNumbered = False
            Exit Function
        End If
        If dic.Exists(k) Then
            CheckAllNodeNumbered = False
            Exit Function
        End If
        dic.Add k, ov
    Next
    CheckAllNodeNumbered = True
End Function

Private Function CheckDisconnection() As Boolean
    Dim sh As Shape
    For Each sh In DrawSheet.Shapes
        ' This magic number -2 is just taken from an inspection result.
        ' It's not yet logically confirmed that -2 always indicate the connector in this usage.
        If sh.Type = msoAutoShape And sh.AutoShapeType = -2 Then
            If sh.ConnectorFormat.BeginConnected And sh.ConnectorFormat.EndConnected Then
                sh.Line.ForeColor.RGB = CONNECTOR_COLOR
            Else
                sh.Line.ForeColor.RGB = vbRed
                CheckDisconnection = False
                Exit Function
            End If
        End If
    Next
    CheckDisconnection = True
End Function

Private Function CheckAllShapeNodeExists() As Boolean
    Let CheckAllShapeNodeExists = False
    With ScheduleSheet
        'Last Used Row Check
        If ScheduleSheet.LastUsedRow < 4 Then
            Exit Function
        End If

        'Count Check
        Dim r As Range
        Dim rr As Range: Set rr = .Range("G4:G" & .LastUsedRow)
        If Not DrawSheet.Ovals.Count = rr.Count Then
            Exit Function
        End If
        
        'Exist Check
        For Each r In rr
            If r.Value = "" Then
                Exit Function
            End If
            On Error GoTo Error_Handler
                Call DrawSheet.Ovals(r.Value)
            On Error GoTo 0
        Next
    End With
    CheckAllShapeNodeExists = True
Exit Function
Error_Handler:
    CheckAllShapeNodeExists = False
End Function

Private Function CheckTraceability(n As Node, node_stack As Nodes, depth As Long) As Boolean
    If depth > DrawSheet.Ovals.Count Then
        CheckTraceability = False
        Exit Function
    End If
    If Not node_stack.Exists(n.ShapeObjectName) Then
        node_stack.AddNode n, n.ShapeObjectName
    End If
    If n.GetDependency.Count > 0 Then
        Dim nn As Node
        For Each nn In n.GetDependency
            Let CheckTraceability = CheckTraceability(nn, node_stack, depth + 1)
            If Not CheckTraceability Then
                Exit Function
            End If
        Next
    Else
        CheckTraceability = n.UnnumberedTaskTitle = "START"
    End If
End Function

Public Sub Btn_PlotSchedule()
    If ConfigSheet.LockMacro Then
        MsgBox "このマクロは既存のデザインに影響を及ぼす可能性があるため、現在ロックされています。" & vbNewLine & "リスクを承知のうえでロックを解除するにはConfigシートのC4セルをFalseに書き換えてください。", vbExclamation
        Exit Sub
    End If
    
    If Not CheckAllNodeNumbered Then
        MsgBox "タスク番号の重複または未設定があります。" & vbCrLf & "Drawシートを確認してください。", vbExclamation, "エラー"
        Exit Sub
    End If

    If Not CheckDisconnection Then
        MsgBox "切断されたコネクターがあります。" & vbCrLf & "Drawシートを確認してください。", vbExclamation, "エラー"
        Exit Sub
    End If

    If Not CheckAllShapeNodeExists Then
        MsgBox "描画されていないタスクまたは余分に描画されたタスクがあります。" & vbCrLf & "Drawシートを確認し、Plot Tasksを実行してください。", vbExclamation, "エラー"
        Exit Sub
    End If
    
    ScheduleSheet.ClearSchedule
    Dim n As Node
    
    Application.Calculation = xlCalculationManual
    Dim nn As Nodes: Set nn = ReadDependency
    
    Dim nnn As Nodes: Set nnn = New Nodes
    If Not CheckTraceability(nn.FindEndNode, nnn, 0) Then
        MsgBox "一部のノードが切断または循環参照されています。" & vbCrLf & "Drawシートを確認してください。", vbExclamation, "エラー"
        Exit Sub
    End If

    If nnn.Count <> DrawSheet.Ovals.Count Then
        MsgBox "一部のノードが切断または循環参照されています。" & vbCrLf & "Drawシートを確認してください。", vbExclamation, "エラー"
        Exit Sub
    End If
    
    For Each n In nn
        Dim tmpRow As Long
        tmpRow = ScheduleSheet.FindRowByShapeName(n.ShapeObjectName)
        
        ScheduleSheet.Cells(tmpRow, ColOffset.Number + 1).Value = n.TaskNumber
        ScheduleSheet.Cells(tmpRow, ColOffset.TaskName + 1).Value = n.UnnumberedTaskTitle
        ScheduleSheet.Cells(tmpRow, ColOffset.PlannedStartDay + 1).NumberFormatLocal = "yyyy/m/d (aaa)"
        ScheduleSheet.Cells(tmpRow, ColOffset.PlannedEndDay + 1).NumberFormatLocal = "yyyy/m/d (aaa)"
        ScheduleSheet.Cells(tmpRow, ColOffset.PlannedEndDay + 1).FormulaR1C1 = "=IF(ISBLANK(RC[4]),WORKDAY(RC[-1],RC[-2],Holidays!C[-5]),RC[-1]+7-WEEKDAY(RC[-1]+7-RC[4]))"
        
        Dim tmpStr As String: tmpStr = ""
        
        Dim n2 As Node
        For Each n2 In n.GetDependency
            Dim tmpRow2 As String: tmpRow2 = ScheduleSheet.FindRowByShapeName(n2.ShapeObjectName)
            tmpStr = tmpStr & "F" & tmpRow2 & ","
        Next
            
        If Len(tmpStr) > 0 Then
            tmpStr = Left(tmpStr, Len(tmpStr) - 1)
        End If
        
        Dim tmpStartOffsetCell As String
        tmpStartOffsetCell = "I" & tmpRow
        
        If n.GetDependency.Count > 0 Then
            ScheduleSheet.Cells(tmpRow, ColOffset.PlannedStartDay + 1).Formula = "=WORKDAY(MAX(" & tmpStr & ")," & tmpStartOffsetCell & " ,Holidays!A:A)"
        End If
        
        tmpStr = ""
        For Each n2 In n.GetDependency
            tmpStr = tmpStr & val(n2.TaskTitle) & ","
        Next
        If Len(tmpStr) > 0 Then tmpStr = Left(tmpStr, Len(tmpStr) - 1)
        ScheduleSheet.Cells(tmpRow, ColOffset.Dependency + 1).Value = tmpStr

        If n.UnnumberedTaskTitle = "START" Then
            ScheduleSheet.Cells(tmpRow, ColOffset.PlannedStartDay + 1).Value = Int(Now())
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
End Sub


Private Function ReadDependency() As Nodes
    'Requre Refference for Microsoft Scripting Runtime Library
    Dim c As Nodes
    Set c = New Nodes
    
    Dim sh As Shape
    For Each sh In DrawSheet.Shapes
        If sh.Type = msoAutoShape And sh.AutoShapeType = 9 Then
            With New Node
                .TaskTitle = Replace(sh.TextFrame2.TextRange.Text, vbLf, "")
                .ShapeObjectName = sh.Name
                c.AddNode .Self, .TaskTitle
            End With
        End If
    Next

    Dim n As Node
    For Each sh In DrawSheet.Shapes
        ' This magic number -2 is just taken from an inspection result.
        ' It's not yet logically confirmed that -2 always indicate the connector in this usage.
        If sh.Type = msoAutoShape And sh.AutoShapeType = -2 Then
            Dim bcs As String
            bcs = Replace(sh.ConnectorFormat.BeginConnectedShape.TextFrame2.TextRange.Text, vbLf, "")
            Dim ecs As String
            ecs = Replace(sh.ConnectorFormat.EndConnectedShape.TextFrame2.TextRange.Text, vbLf, "")
            
            Set n = c.Item(ecs)
            n.AddDependency c.Item(bcs)
        End If
    Next
    
    c.Sort

    Set ReadDependency = c
End Function

Public Sub Btn_FillDefault()
    If ConfigSheet.LockMacro Then
        MsgBox "このマクロは既存のデザインに影響を及ぼす可能性があるため、現在ロックされています。" & vbNewLine & "リスクを承知のうえでロックを解除するにはConfigシートのC4セルをFalseに書き換えてください。", vbExclamation
        Exit Sub
    End If
    Dim i As Long
    Application.Calculation = xlCalculationManual
    For i = 4 To ScheduleSheet.LastUsedRow
        If ScheduleSheet.Range("B" & i).Value = "START" Or _
            ScheduleSheet.Range("B" & i).Value = "END" Then
            ScheduleSheet.Range("D" & i).Value = 0
            ScheduleSheet.Range("I" & i).Value = 0
        Else
            ScheduleSheet.Range("D" & i).Value = 1
            If ScheduleSheet.Range("C" & i).Value = 0 Then
                ScheduleSheet.Range("I" & i).Value = 0
            Else
                ScheduleSheet.Range("I" & i).Value = 1
            End If
        End If
    Next
    Application.Calculate
    Application.Calculation = xlCalculationAutomatic
End Sub

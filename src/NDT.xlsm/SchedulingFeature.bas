Attribute VB_Name = "SchedulingFeature"
Option Explicit
Const CONNECTOR_COLOR = XlRgbColor.rgbDimGray
Function CheckAllNodeNumbered() As Boolean
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

Function CheckDisconnection() As Boolean
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

Sub PlotSchedule()
    If Not CheckAllNodeNumbered Then
        MsgBox "タスク番号の重複または未設定があります。" & vbCrLf & "Drawシートを確認してください。", vbExclamation, "エラー"
        Exit Sub
    End If

    If Not CheckDisconnection Then
        MsgBox "切断されたコネクターがあります。" & vbCrLf & "Drawシートを確認してください。", vbExclamation, "エラー"
        Exit Sub
    End If

    ScheduleSheet.ClearSchedule
    Dim n As Node
    
    Application.Calculation = xlCalculationManual
    For Each n In ReadDependency
        Dim tmpRow As Long
        tmpRow = ScheduleSheet.FindRowByShapeName(n.ShapeObjectName)
        
        ScheduleSheet.Cells(tmpRow, ColOffset.Number + 1).Value = n.TaskNumber
        ScheduleSheet.Cells(tmpRow, ColOffset.TaskName + 1).Value = n.UnnumberedTaskTitle
        ScheduleSheet.Cells(tmpRow, ColOffset.PlannedStartDay + 1).NumberFormatLocal = "yyyy/m/d"
        ScheduleSheet.Cells(tmpRow, ColOffset.PlannedEndDay + 1).NumberFormatLocal = "yyyy/m/d"
        ScheduleSheet.Cells(tmpRow, ColOffset.PlannedEndDay + 1).FormulaR1C1 = "=WORKDAY(RC[-1],RC[-2],Holidays!C[-5])"
        
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
            tmpStr = tmpStr & Val(n2.TaskTitle) & ","
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

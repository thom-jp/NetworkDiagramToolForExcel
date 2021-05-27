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

    ScheduleSheet.ClearAllData
    Dim n As Node
    
    Application.Calculation = xlCalculationManual
    For Each n In ReadDependency
        With ScheduleSheet.DataStartCell
            .Offset(n.TaskNumber, ColOffset.Number).Value = n.TaskNumber
            .Offset(n.TaskNumber, ColOffset.TaskName).Value = n.UnnumberedTaskTitle
            .Offset(n.TaskNumber, ColOffset.PlannedStartDay).NumberFormatLocal = "yyyy/m/d"
            .Offset(n.TaskNumber, ColOffset.PlannedEndDay).NumberFormatLocal = "yyyy/m/d"
            .Offset(n.TaskNumber, ColOffset.PlannedEndDay).FormulaR1C1 = "=WORKDAY(RC[-1],RC[-2],Holidays!C[-5])"
            .Offset(n.TaskNumber, ColOffset.Duration).Value = 1
        End With
        
        Dim tmpStr As String: tmpStr = ""
        
        Dim n2 As Node
        For Each n2 In n.GetDependency
            tmpStr = tmpStr & "F" & n2.TaskNumber + ScheduleSheet.DataStartCell.Row & ","
        Next
            
        If Len(tmpStr) > 0 Then
            tmpStr = Left(tmpStr, Len(tmpStr) - 1)
        End If
        
        If n.GetDependency.Count > 0 Then
            ScheduleSheet.DataStartCell.Offset(n.TaskNumber, ColOffset.PlannedStartDay).Formula = "=WORKDAY(MAX(" & tmpStr & "),1,Holidays!A:A)"
        End If
        
        tmpStr = ""
        For Each n2 In n.GetDependency
            tmpStr = tmpStr & Val(n2.TaskTitle) & ","
        Next
        If Len(tmpStr) > 0 Then tmpStr = Left(tmpStr, Len(tmpStr) - 1)
        ScheduleSheet.DataStartCell.Offset(n.TaskNumber, ColOffset.Dependency).Value = tmpStr
    Next
    ScheduleSheet.DataStartCell.Offset(0, ColOffset.PlannedStartDay).Value = Int(Now())
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

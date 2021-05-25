Attribute VB_Name = "SchedulingFeature"
Option Explicit
Sub PlotSchedule()
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


Function ReadDependency() As Collection
    'Requre Refference for Microsoft Scripting Runtime Library
    Dim nodes As Scripting.Dictionary
    Set nodes = New Scripting.Dictionary
    
    Dim sh As Shape
    For Each sh In DrawSheet.Shapes
        If sh.Type = msoAutoShape And sh.AutoShapeType = 9 Then
            With New Node
                .TaskTitle = Replace(sh.TextFrame2.TextRange.Text, vbLf, "")
                nodes.Add .TaskTitle, .Self
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
            
            Set n = nodes.Item(ecs)
            n.AddDependency nodes.Item(bcs)
        End If
    Next
    
    Dim c As Collection: Set c = New Collection
    Dim k As Variant
    For Each k In nodes.Keys
        c.Add nodes.Item(k)
    Next
    
    CSort c, "SortKey1"

    Set ReadDependency = c
End Function

Attribute VB_Name = "SchedulingFeature"
Sub PlotSchedule()
    Dim c As Collection
    Set c = ReadDependency
    Dim n As Node
    
    Application.Calculation = xlCalculationManual
    For Each n In c
        ScheduleSheet.Range("B4").Offset(Val(n.TaskTitle), 0).Value = n.TaskTitle
        ScheduleSheet.Range("E4").Offset(Val(n.TaskTitle), 0).NumberFormatLocal = "yyyy/m/d"
        ScheduleSheet.Range("F4").Offset(Val(n.TaskTitle), 0).NumberFormatLocal = "yyyy/m/d"
        ScheduleSheet.Range("D4").Offset(Val(n.TaskTitle), 0).Value = 1
        ScheduleSheet.Range("F4").Offset(Val(n.TaskTitle), 0).FormulaR1C1 = "=WORKDAY(RC[-1],RC[-2],Holidays!C[-5])"
        tmpStr = ""
        Dim n2 As Node
        For Each n2 In n.GetDependency
            tmpStr = tmpStr & "F" & Val(n2.TaskTitle) + 4 & ","
        Next
            
        If Len(tmpStr) > 0 Then tmpStr = Left(tmpStr, Len(tmpStr) - 1)
        If n.GetDependency.Count > 0 Then ScheduleSheet.Range("E4").Offset(Val(n.TaskTitle), 0).Formula = "=WORKDAY(MAX(" & tmpStr & "),1,Holidays!A:A)"
        
        tmpStr = ""
        For Each n2 In n.GetDependency
            tmpStr = tmpStr & Val(n2.TaskTitle) & ","
        Next
        If Len(tmpStr) > 0 Then tmpStr = Left(tmpStr, Len(tmpStr) - 1)
        ScheduleSheet.Range("C4").Offset(Val(n.TaskTitle), 0).Value = tmpStr
    Next
    Range("E4").Value = Int(Now())
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
            
            bcs = Replace(sh.ConnectorFormat.BeginConnectedShape.TextFrame2.TextRange.Text, vbLf, "")
            ecs = Replace(sh.ConnectorFormat.EndConnectedShape.TextFrame2.TextRange.Text, vbLf, "")
            
            Set n = nodes.Item(ecs)
            n.AddDependency nodes.Item(bcs)
        End If
    Next
    
    Dim c As Collection: Set c = New Collection
    For Each k In nodes.Keys
        c.Add nodes.Item(k)
    Next
    
    CSort c, "SortKey1"

    Set ReadDependency = c
End Function

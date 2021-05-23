Attribute VB_Name = "SchedulingFeature"
Sub ReadDependency()
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
    
    For Each k In nodes.Keys
        Set n = nodes.Item(k)
        n.DumpStatus
    Next
End Sub

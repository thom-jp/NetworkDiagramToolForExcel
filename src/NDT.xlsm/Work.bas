Attribute VB_Name = "Work"
Option Explicit
Sub TestFeature()
    Application.ScreenUpdating = False
    DrawSheet.Select
    DrawFeature.RemoveAllShapse
    TaskListSheet.Select
    TaskListSheet.Range(Range("A4"), Range("A4").End(xlDown)).Select
    DrawFeature.DrawTaskAsNode
    DrawSheet.Select
    Application.ScreenUpdating = True
End Sub

Sub TestNodesCollection()
    Dim nn As Nodes
    Set nn = New Nodes
    
    Dim r As Range
    For Each r In Selection
        With New Node
            .TaskTitle = r.Value
            nn.AddNode .Self
        End With
    Next
    
    Dim n As Node
    For Each n In nn
        Debug.Print n.TaskTitle
    Next
End Sub

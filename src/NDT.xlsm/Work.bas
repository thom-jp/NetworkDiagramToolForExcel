Attribute VB_Name = "Work"
Option Explicit

Sub TestNodesCollection()
    Dim nn As Nodes
    Set nn = New Nodes
    
    Dim r As Range
    Dim i As Integer
    For Each r In Selection
        With New Node
            .TaskTitle = WorksheetFunction.RandBetween(1, 10000) & "." & r.Value
            nn.AddNode .Self, CStr(i)
            i = i + 1
        End With
    Next
    
    Debug.Print nn.Exists("1")
    Debug.Print nn.Exists("10000")
    
    Dim n As Node
    For Each n In nn
        Debug.Print n.TaskTitle
    Next
    
    Debug.Print String(10, "-")
    nn.Sort
    For Each n In nn
        Debug.Print n.TaskTitle
    Next
End Sub

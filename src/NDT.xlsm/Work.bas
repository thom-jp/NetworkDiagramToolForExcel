Attribute VB_Name = "Work"
Sub TestFeature()
    Application.ScreenUpdating = False
    DrawSheet.Select
    DrawFeature.RemoveAllShapse
    DataSheet.Select
    DataSheet.Range(Range("c4"), Range("c4").End(xlDown)).Select
    DrawFeature.DrawTaskAsNode
    DrawSheet.Select
    Application.ScreenUpdating = True
End Sub

Sub testNodeClassBehavior()
    Dim n As Node
    Set n = New Node
    n.TaskTitle = "Hoge"
    
    Dim n2 As Node
    Set n2 = New Node
    n2.TaskTitle = "Fuga"
    
    Dim n3 As Node
    Set n3 = New Node
    n3.TaskTitle = "Piyo"
    
    n.AddDependency n2
    n.AddDependency n3
    
    n.DumpStatus
    
End Sub

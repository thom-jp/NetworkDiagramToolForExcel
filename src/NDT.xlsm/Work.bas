Attribute VB_Name = "Work"
Sub TestFeature()
    Application.ScreenUpdating = False
    DrawSheet.Select
    Module1.RemoveAllShapse
    DataSheet.Select
    DataSheet.Range(Range("c4"), Range("c4").End(xlDown)).Select
    Module1.DrawTaskAsNode
    DrawSheet.Select
    Application.ScreenUpdating = True
End Sub

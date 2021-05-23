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

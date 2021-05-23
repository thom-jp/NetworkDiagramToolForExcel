Attribute VB_Name = "Work"
Sub TestFeature()
    DrawSheet.Select
    Module1.RemoveAllShapse
    DataSheet.Select
    DataSheet.Range(Range("c4"), Range("c4").End(xlDown)).Select
    Module1.DrawTaskAsNode
    DrawSheet.Select
End Sub

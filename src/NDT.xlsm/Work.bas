Attribute VB_Name = "Work"
Sub TestFeature()
    DrawSheet.Select
    Module1.RemoveAllShapse
    DataSheet.Select
    DataSheet.Range("c4:c12").Select
    Module1.DrawTaskAsNode
    DrawSheet.Select
End Sub

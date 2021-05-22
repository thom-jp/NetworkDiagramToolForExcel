Attribute VB_Name = "Module1"
Function CreateNodeShape(size, pos_x, pos_y, task_title As String) As Shape
    Dim s As Shape
    Set s = DrawSheet.Shapes.AddShape(msoShapeOval, pos_x, pos_y, size, size)
    s.TextFrame2.TextRange.Characters.Text = task_title
End Function

Sub DrawTaskAsNode()
    size = 60
    x = 10: y = 10
    Dim r As Range
    For Each r In Selection
        Call CreateNodeShape(size, x, y, r.Value)
        x = x + size + 10
    Next
End Sub

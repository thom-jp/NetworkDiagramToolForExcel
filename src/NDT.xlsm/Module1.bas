Attribute VB_Name = "Module1"
Const SIZE = 60
Const Y_DISTANCE = 20
Const X_DISTANCE = 10
Function CreateNodeShape(SIZE, pos_x, pos_y, task_title As String) As Shape
    Dim s As Shape
    Set s = DrawSheet.Shapes.AddShape(msoShapeOval, pos_x, pos_y, SIZE, SIZE)
    s.TextFrame2.TextRange.Text = task_title
End Function

Sub DrawTaskAsNode()
    x = X_DISTANCE: y = Y_DISTANCE
    Dim r As Range
    For Each r In Selection
        Call CreateNodeShape(SIZE, x, y, r.Value)
        x = x + SIZE + X_DISTANCE
    Next
End Sub

Sub OrderNodeVertical()
    'To keep selection order, store shapes to a Collection.
    Dim c As Collection
    Set c = New Collection
    Dim sh As Shape
    For Each sh In Selection.ShapeRange
        c.Add sh
    Next

    Selection.ShapeRange.Group.Select
    leftEdge = Round(Selection.ShapeRange.Left, 3)
    topEdge = Selection.ShapeRange.Top
    Selection.ShapeRange.Ungroup
    
    For Each sh In c
        sh.Left = leftEdge
        sh.Top = topEdge + (SIZE + Y_DISTANCE) * n
        n = n + 1
    Next
End Sub


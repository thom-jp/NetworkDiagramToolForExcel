Attribute VB_Name = "Module1"
Sub Macro1()
    pos_x = 0: pos_y = 0: Size = 60
    TaskTitle = "START"
    
    Dim s As Shape
    Set s = DrawSheet.Shapes.AddShape(msoShapeOval, pos_x, pos_y, Size, Size)
    s.TextFrame2.TextRange.Characters.Text = TaskTitle
End Sub

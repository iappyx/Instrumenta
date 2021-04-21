Attribute VB_Name = "ModuleObjectsRoundedCorners"
Sub ObjectsSetRoundedCorner(ShapeRadius As Single)
    Dim SlideShape  As PowerPoint.Shape
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        With SlideShape
            .AutoShapeType = msoShapeRoundedRectangle
            .Adjustments(1) = (1 / (SlideShape.Height + SlideShape.Width)) * ShapeRadius
        End With
    Next
End Sub

Sub ObjectsCopyRoundedCorner()
    Dim SlideShape  As PowerPoint.Shape
    Set myDocument = Application.ActiveWindow
    Dim ShapeRadius As Single
    ShapeRadius = myDocument.Selection.ShapeRange(1).Adjustments(1) / (1 / (myDocument.Selection.ShapeRange(1).Height + myDocument.Selection.ShapeRange(1).Width))
    ObjectsSetRoundedCorner (ShapeRadius)
End Sub

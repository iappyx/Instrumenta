Attribute VB_Name = "ModuleObjectsRoundedCorners"
Sub ObjectsCopyRoundedCorner()
    Dim SlideShape  As PowerPoint.Shape
    Set myDocument = Application.ActiveWindow
    Dim ShapeRadius As Single
    ShapeRadius = myDocument.Selection.ShapeRange(1).Adjustments(1) / (1 / (myDocument.Selection.ShapeRange(1).Height + myDocument.Selection.ShapeRange(1).Width))
    
    If myDocument.Selection.ShapeRange(1).Adjustments.Count > 1 Then
        ShapeRadius2 = myDocument.Selection.ShapeRange(1).Adjustments(2) / (1 / (myDocument.Selection.ShapeRange(1).Height + myDocument.Selection.ShapeRange(1).Width))
    End If
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        With SlideShape
            .AutoShapeType = myDocument.Selection.ShapeRange(1).AutoShapeType
            .Adjustments(1) = (1 / (SlideShape.Height + SlideShape.Width)) * ShapeRadius
            If myDocument.Selection.ShapeRange(1).Adjustments.Count > 1 Then
                .Adjustments(2) = (1 / (SlideShape.Height + SlideShape.Width)) * ShapeRadius2
            End If
        End With
    Next
    
End Sub

Sub ObjectsCopyShapeTypeAndAdjustments()
    Dim SlideShape  As PowerPoint.Shape
    Set myDocument = Application.ActiveWindow
    Dim AdjustmentsCount As Long
    Dim ShapeCount  As Long
    
    For ShapeCount = 2 To ActiveWindow.Selection.ShapeRange.Count
        
        myDocument.Selection.ShapeRange(ShapeCount).AutoShapeType = myDocument.Selection.ShapeRange(1).AutoShapeType
        
        For AdjustmentsCount = 1 To myDocument.Selection.ShapeRange(1).Adjustments.Count
            
            myDocument.Selection.ShapeRange(ShapeCount).Adjustments(AdjustmentsCount) = myDocument.Selection.ShapeRange(1).Adjustments(AdjustmentsCount)
            
        Next AdjustmentsCount
        
    Next ShapeCount
    
End Sub

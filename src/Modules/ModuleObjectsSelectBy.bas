Attribute VB_Name = "ModuleObjectsSelectBy"
Sub ObjectsSelectBySameFillAndLineColor()
    
    Set myDocument = Application.ActiveWindow
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Fill.ForeColor = SlideShape.Fill.ForeColor) And (SlideShapeToCheck.Line.ForeColor = SlideShape.Line.ForeColor) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.Visible = True) And (SlideShapeToCheck.Line.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
End Sub

Sub ObjectsSelectBySameFillColor()
    
    Set myDocument = Application.ActiveWindow
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Fill.ForeColor = SlideShape.Fill.ForeColor) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
End Sub

Sub ObjectsSelectBySameLineColor()
    
    Set myDocument = Application.ActiveWindow
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Line.ForeColor = SlideShape.Line.ForeColor) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Line.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
End Sub

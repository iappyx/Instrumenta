Attribute VB_Name = "ModuleObjectsAlignAndDistribute"
Sub ObjectsAlignLefts()
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange.Count = 1 Then
        myDocument.Selection.ShapeRange.Align msoAlignLefts, msoTrue
    Else
        myDocument.Selection.ShapeRange.Align msoAlignLefts, msoFalse
    End If
    
End Sub

Sub ObjectsAlignRights()
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange.Count = 1 Then
        myDocument.Selection.ShapeRange.Align msoAlignRights, msoTrue
    Else
        myDocument.Selection.ShapeRange.Align msoAlignRights, msoFalse
    End If
    
End Sub

Sub ObjectsAlignBottoms()
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange.Count = 1 Then
        myDocument.Selection.ShapeRange.Align msoAlignBottoms, msoTrue
    Else
        myDocument.Selection.ShapeRange.Align msoAlignBottoms, msoFalse
    End If
    
End Sub

Sub ObjectsAlignCenters()
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange.Count = 1 Then
        myDocument.Selection.ShapeRange.Align msoAlignCenters, msoTrue
    Else
        myDocument.Selection.ShapeRange.Align msoAlignCenters, msoFalse
    End If
    
End Sub

Sub ObjectsAlignMiddles()
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange.Count = 1 Then
        myDocument.Selection.ShapeRange.Align msoAlignMiddles, msoTrue
    Else
        myDocument.Selection.ShapeRange.Align msoAlignMiddles, msoFalse
    End If
    
End Sub

Sub ObjectsAlignTops()
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange.Count = 1 Then
        myDocument.Selection.ShapeRange.Align msoAlignTops, msoTrue
    Else
        myDocument.Selection.ShapeRange.Align msoAlignTops, msoFalse
    End If
    
End Sub

Sub ObjectsDistributeHorizontally()
    Set myDocument = Application.ActiveWindow
    
    myDocument.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse
    
End Sub

Sub ObjectsDistributeVertically()
    Set myDocument = Application.ActiveWindow
    
    myDocument.Selection.ShapeRange.Distribute msoDistributeVertically, msoFalse
    
End Sub

Attribute VB_Name = "ModuleObjectsAlignAndDistribute"
Sub ObjectsRemoveSpacingHorizontal()
    
    Set myDocument = Application.ActiveWindow
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
    
    For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
        Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
    Next ShapeCount
    
    ObjectsSortByLeftPosition SlideShape
    
    For ShapeCount = 2 To UBound(SlideShape)
        SlideShape(ShapeCount).Left = SlideShape(ShapeCount - 1).Left + SlideShape(ShapeCount - 1).Width
    Next ShapeCount
End Sub

Sub ObjectsRemoveSpacingVertical()
    
    Set myDocument = Application.ActiveWindow
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
    
    For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
        Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
    Next ShapeCount
    
    ObjectsSortByTopPosition SlideShape
    
    For ShapeCount = 2 To UBound(SlideShape)
        SlideShape(ShapeCount).Top = SlideShape(ShapeCount - 1).Top + SlideShape(ShapeCount - 1).Height
    Next ShapeCount
End Sub

Sub ObjectsIncreaseSpacingHorizontal()
    
    Set myDocument = Application.ActiveWindow
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
    
    For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
        Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
    Next ShapeCount
    
    ObjectsSortByLeftPosition SlideShape
    
    For ShapeCount = 2 To UBound(SlideShape)
        SlideShape(ShapeCount).Left = SlideShape(ShapeCount).Left + (ShapeCount - 1) * 0.2
    Next ShapeCount
End Sub

Sub ObjectsDecreaseSpacingHorizontal()
    
    Set myDocument = Application.ActiveWindow
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
    
    For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
        Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
    Next ShapeCount
    
    ObjectsSortByLeftPosition SlideShape
    
    For ShapeCount = 2 To UBound(SlideShape)
        SlideShape(ShapeCount).Left = SlideShape(ShapeCount).Left - (ShapeCount - 1) * 0.2
    Next ShapeCount
End Sub

Sub ObjectsIncreaseSpacingVertical()
    
    Set myDocument = Application.ActiveWindow
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
    
    For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
        Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
    Next ShapeCount
    
    ObjectsSortByTopPosition SlideShape
    
    For ShapeCount = 2 To UBound(SlideShape)
        SlideShape(ShapeCount).Top = SlideShape(ShapeCount).Top + (ShapeCount - 1) * 0.2
    Next ShapeCount
End Sub

Sub ObjectsDecreaseSpacingVertical()
    
    Set myDocument = Application.ActiveWindow
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
    
    For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
        Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
    Next ShapeCount
    
    ObjectsSortByTopPosition SlideShape
    
    For ShapeCount = 2 To UBound(SlideShape)
        SlideShape(ShapeCount).Top = SlideShape(ShapeCount).Top - (ShapeCount - 1) * 0.2
    Next ShapeCount
End Sub

Sub ObjectsSortByLeftPosition(ArrayToSort As Variant)
    
    Dim StopLoop    As Boolean
    Dim ShapeCount  As Long
    Dim SlideShapes As Shape
    Do
        StopLoop = False
        For ShapeCount = LBound(ArrayToSort) To UBound(ArrayToSort) - 1
            
            If ArrayToSort(ShapeCount).Left > ArrayToSort(ShapeCount + 1).Left Then
                Set SlideShapes = ArrayToSort(ShapeCount)
                Set ArrayToSort(ShapeCount) = ArrayToSort(ShapeCount + 1)
                Set ArrayToSort(ShapeCount + 1) = SlideShapes
                StopLoop = True
            End If
        Next ShapeCount
    Loop Until Not StopLoop
    
    Set SlideShapes = Nothing
End Sub

Sub ObjectsSortByTopPosition(ArrayToSort As Variant)
    
    Dim StopLoop    As Boolean
    Dim ShapeCount  As Long
    Dim SlideShapes As Shape
    Do
        StopLoop = False
        For ShapeCount = LBound(ArrayToSort) To UBound(ArrayToSort) - 1
            
            If ArrayToSort(ShapeCount).Top > ArrayToSort(ShapeCount + 1).Top Then
                Set SlideShapes = ArrayToSort(ShapeCount)
                Set ArrayToSort(ShapeCount) = ArrayToSort(ShapeCount + 1)
                Set ArrayToSort(ShapeCount + 1) = SlideShapes
                StopLoop = True
            End If
        Next ShapeCount
    Loop Until Not StopLoop
    
    Set SlideShapes = Nothing
End Sub

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

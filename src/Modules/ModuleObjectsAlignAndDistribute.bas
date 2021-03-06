Attribute VB_Name = "ModuleObjectsAlignAndDistribute"
'MIT License

'Copyright (c) 2021 iappyx

'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Sub ObjectsStretchTop()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    
    If myDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To myDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Height = SlideShape(ShapeCount).Height + (SlideShape(ShapeCount).Top - SlideShape(1).Top)
            SlideShape(ShapeCount).Top = SlideShape(1).Top
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Height = SlideShape(ShapeCount).Height + (SlideShape(ShapeCount).Top - SlideShape(1).Top)
            SlideShape(ShapeCount).Top = SlideShape(1).Top
        Next ShapeCount
        
    End If
    
End Sub

Sub ObjectsStretchLeft()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    
    If myDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To myDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Width = SlideShape(ShapeCount).Width + (SlideShape(ShapeCount).Left - SlideShape(1).Left)
            SlideShape(ShapeCount).Left = SlideShape(1).Left
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Width = SlideShape(ShapeCount).Width + (SlideShape(ShapeCount).Left - SlideShape(1).Left)
            SlideShape(ShapeCount).Left = SlideShape(1).Left
        Next ShapeCount
        
    End If
    
End Sub

Sub ObjectsStretchBottom()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    
    If myDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To myDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByBottomPosition SlideShape
        
        For ShapeCount = UBound(SlideShape) - 1 To 1 Step -1
            SlideShape(ShapeCount).Height = SlideShape(ShapeCount).Height + ((SlideShape(UBound(SlideShape)).Top + SlideShape(UBound(SlideShape)).Height) - SlideShape(ShapeCount).Top - SlideShape(ShapeCount).Height)
            
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByBottomPosition SlideShape
        
        For ShapeCount = UBound(SlideShape) - 1 To 1 Step -1
            SlideShape(ShapeCount).Height = SlideShape(ShapeCount).Height + ((SlideShape(UBound(SlideShape)).Top + SlideShape(UBound(SlideShape)).Height) - SlideShape(ShapeCount).Top - SlideShape(ShapeCount).Height)
            
        Next ShapeCount
        
    End If
    
End Sub

Sub ObjectsStretchRight()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    
    If myDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To myDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByRightPosition SlideShape
        
        For ShapeCount = UBound(SlideShape) - 1 To 1 Step -1
            SlideShape(ShapeCount).Width = SlideShape(ShapeCount).Width + ((SlideShape(UBound(SlideShape)).Left + SlideShape(UBound(SlideShape)).Width) - SlideShape(ShapeCount).Left - SlideShape(ShapeCount).Width)
            
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByRightPosition SlideShape
        
        For ShapeCount = UBound(SlideShape) - 1 To 1 Step -1
            SlideShape(ShapeCount).Width = SlideShape(ShapeCount).Width + ((SlideShape(UBound(SlideShape)).Left + SlideShape(UBound(SlideShape)).Width) - SlideShape(ShapeCount).Left - SlideShape(ShapeCount).Width)
            
        Next ShapeCount
        
    End If
End Sub

Sub ObjectsRemoveSpacingHorizontal()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    
    If myDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To myDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Left = SlideShape(ShapeCount - 1).Left + SlideShape(ShapeCount - 1).Width
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Left = SlideShape(ShapeCount - 1).Left + SlideShape(ShapeCount - 1).Width
        Next ShapeCount
        
    End If
End Sub

Sub ObjectsRemoveSpacingVertical()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    
    If myDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To myDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Top = SlideShape(ShapeCount - 1).Top + SlideShape(ShapeCount - 1).Height
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Top = SlideShape(ShapeCount - 1).Top + SlideShape(ShapeCount - 1).Height
        Next ShapeCount
        
    End If
End Sub

Sub ObjectsIncreaseSpacingHorizontal()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    
    If myDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To myDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Left = SlideShape(ShapeCount).Left + (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    Else
        ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Left = SlideShape(ShapeCount).Left + (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    End If
    
End Sub

Sub ObjectsDecreaseSpacingHorizontal()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    
    If myDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To myDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Left = SlideShape(ShapeCount).Left - (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Left = SlideShape(ShapeCount).Left - (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    End If
End Sub

Sub ObjectsIncreaseSpacingVertical()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    
    If myDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To myDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Top = SlideShape(ShapeCount).Top + (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Top = SlideShape(ShapeCount).Top + (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    End If
    
End Sub

Sub ObjectsDecreaseSpacingVertical()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As Shape
    
    If myDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To myDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Top = SlideShape(ShapeCount).Top - (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Top = SlideShape(ShapeCount).Top - (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    End If
    
End Sub

Sub ObjectsSortByLeftPosition(ArrayToSort As Variant)
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
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

Sub ObjectsSortByRightPosition(ArrayToSort As Variant)
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim StopLoop    As Boolean
    Dim ShapeCount  As Long
    Dim SlideShapes As Shape
    Do
        StopLoop = False
        For ShapeCount = LBound(ArrayToSort) To UBound(ArrayToSort) - 1
            
            If (ArrayToSort(ShapeCount).Left + ArrayToSort(ShapeCount).Width) > (ArrayToSort(ShapeCount + 1).Left + ArrayToSort(ShapeCount + 1).Width) Then
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
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
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

Sub ObjectsSortByBottomPosition(ArrayToSort As Variant)
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim StopLoop    As Boolean
    Dim ShapeCount  As Long
    Dim SlideShapes As Shape
    Do
        StopLoop = False
        For ShapeCount = LBound(ArrayToSort) To UBound(ArrayToSort) - 1
            
            If (ArrayToSort(ShapeCount).Top + ArrayToSort(ShapeCount).Height) > (ArrayToSort(ShapeCount + 1).Top + ArrayToSort(ShapeCount + 1).Height) Then
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
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.HasChildShapeRange Then
        If myDocument.Selection.ChildShapeRange.Count > 1 Then
            myDocument.Selection.ChildShapeRange.Align msoAlignLefts, msoFalse
        End If
    ElseIf myDocument.Selection.ShapeRange.Count = 1 Then
        myDocument.Selection.ShapeRange.Align msoAlignLefts, msoTrue
    Else
        myDocument.Selection.ShapeRange.Align msoAlignLefts, msoFalse
    End If
    
End Sub

Sub ObjectsAlignRights()
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.HasChildShapeRange Then
        If myDocument.Selection.ChildShapeRange.Count > 1 Then
            myDocument.Selection.ChildShapeRange.Align msoAlignRights, msoFalse
        End If
    ElseIf myDocument.Selection.ShapeRange.Count = 1 Then
        myDocument.Selection.ShapeRange.Align msoAlignRights, msoTrue
    Else
        myDocument.Selection.ShapeRange.Align msoAlignRights, msoFalse
    End If
    
End Sub

Sub ObjectsAlignBottoms()
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.HasChildShapeRange Then
        If myDocument.Selection.ChildShapeRange.Count > 1 Then
            myDocument.Selection.ChildShapeRange.Align msoAlignBottoms, msoFalse
        End If
    ElseIf myDocument.Selection.ShapeRange.Count = 1 Then
        myDocument.Selection.ShapeRange.Align msoAlignBottoms, msoTrue
    Else
        myDocument.Selection.ShapeRange.Align msoAlignBottoms, msoFalse
    End If
    
End Sub

Sub ObjectsAlignCenters()
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.HasChildShapeRange Then
        If myDocument.Selection.ChildShapeRange.Count > 1 Then
            myDocument.Selection.ChildShapeRange.Align msoAlignCenters, msoFalse
        End If
    ElseIf myDocument.Selection.ShapeRange.Count = 1 Then
        myDocument.Selection.ShapeRange.Align msoAlignCenters, msoTrue
    Else
        myDocument.Selection.ShapeRange.Align msoAlignCenters, msoFalse
    End If
    
End Sub

Sub ObjectsAlignMiddles()
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.HasChildShapeRange Then
        If myDocument.Selection.ChildShapeRange.Count > 1 Then
            myDocument.Selection.ChildShapeRange.Align msoAlignMiddles, msoFalse
        End If
    ElseIf myDocument.Selection.ShapeRange.Count = 1 Then
        myDocument.Selection.ShapeRange.Align msoAlignMiddles, msoTrue
    Else
        myDocument.Selection.ShapeRange.Align msoAlignMiddles, msoFalse
    End If
    
End Sub

Sub ObjectsAlignTops()
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.HasChildShapeRange Then
        If myDocument.Selection.ChildShapeRange.Count > 1 Then
            myDocument.Selection.ChildShapeRange.Align msoAlignTops, msoFalse
        End If
    ElseIf myDocument.Selection.ShapeRange.Count = 1 Then
        myDocument.Selection.ShapeRange.Align msoAlignTops, msoTrue
    Else
        myDocument.Selection.ShapeRange.Align msoAlignTops, msoFalse
    End If
    
End Sub

Sub ObjectsDistributeHorizontally()
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.ShapeRange.Count > 2 Then
        myDocument.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse
        
    ElseIf myDocument.Selection.HasChildShapeRange Then
        
        If myDocument.Selection.ChildShapeRange.Count > 2 Then
            myDocument.Selection.ChildShapeRange.Distribute msoDistributeHorizontally, msoFalse
        End If
        
    Else
        MsgBox "Select more shapes to use this command."
    End If
    
End Sub

Sub ObjectsDistributeVertically()
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.ShapeRange.Count > 2 Then
        myDocument.Selection.ShapeRange.Distribute msoDistributeVertically, msoFalse
        
    ElseIf myDocument.Selection.HasChildShapeRange Then
        
        If myDocument.Selection.ChildShapeRange.Count > 2 Then
            
            myDocument.Selection.ChildShapeRange.Distribute msoDistributeVertically, msoFalse
            
        End If
        
    Else
        MsgBox "Select more shapes to use this command."
    End If
End Sub

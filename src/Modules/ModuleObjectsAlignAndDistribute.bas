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
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As shape
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To MyDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Height = SlideShape(ShapeCount).Height + (SlideShape(ShapeCount).Top - SlideShape(1).Top)
            SlideShape(ShapeCount).Top = SlideShape(1).Top
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To MyDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Height = SlideShape(ShapeCount).Height + (SlideShape(ShapeCount).Top - SlideShape(1).Top)
            SlideShape(ShapeCount).Top = SlideShape(1).Top
        Next ShapeCount
        
    End If
    
End Sub

Sub ObjectsStretchLeft()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As shape
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To MyDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Width = SlideShape(ShapeCount).Width + (SlideShape(ShapeCount).left - SlideShape(1).left)
            SlideShape(ShapeCount).left = SlideShape(1).left
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To MyDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Width = SlideShape(ShapeCount).Width + (SlideShape(ShapeCount).left - SlideShape(1).left)
            SlideShape(ShapeCount).left = SlideShape(1).left
        Next ShapeCount
        
    End If
    
End Sub

Sub ObjectsStretchBottom()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As shape
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To MyDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByBottomPosition SlideShape
        
        For ShapeCount = UBound(SlideShape) - 1 To 1 Step -1
            SlideShape(ShapeCount).Height = SlideShape(ShapeCount).Height + ((SlideShape(UBound(SlideShape)).Top + SlideShape(UBound(SlideShape)).Height) - SlideShape(ShapeCount).Top - SlideShape(ShapeCount).Height)
            
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To MyDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByBottomPosition SlideShape
        
        For ShapeCount = UBound(SlideShape) - 1 To 1 Step -1
            SlideShape(ShapeCount).Height = SlideShape(ShapeCount).Height + ((SlideShape(UBound(SlideShape)).Top + SlideShape(UBound(SlideShape)).Height) - SlideShape(ShapeCount).Top - SlideShape(ShapeCount).Height)
            
        Next ShapeCount
        
    End If
    
End Sub

Sub ObjectsStretchRight()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As shape
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To MyDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByRightPosition SlideShape
        
        For ShapeCount = UBound(SlideShape) - 1 To 1 Step -1
            SlideShape(ShapeCount).Width = SlideShape(ShapeCount).Width + ((SlideShape(UBound(SlideShape)).left + SlideShape(UBound(SlideShape)).Width) - SlideShape(ShapeCount).left - SlideShape(ShapeCount).Width)
            
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To MyDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByRightPosition SlideShape
        
        For ShapeCount = UBound(SlideShape) - 1 To 1 Step -1
            SlideShape(ShapeCount).Width = SlideShape(ShapeCount).Width + ((SlideShape(UBound(SlideShape)).left + SlideShape(UBound(SlideShape)).Width) - SlideShape(ShapeCount).left - SlideShape(ShapeCount).Width)
            
        Next ShapeCount
        
    End If
End Sub

Sub ObjectsRemoveSpacingHorizontal()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As shape
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To MyDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).left = SlideShape(ShapeCount - 1).left + SlideShape(ShapeCount - 1).Width
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To MyDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).left = SlideShape(ShapeCount - 1).left + SlideShape(ShapeCount - 1).Width
        Next ShapeCount
        
    End If
End Sub

Sub ObjectsRemoveSpacingVertical()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As shape
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To MyDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Top = SlideShape(ShapeCount - 1).Top + SlideShape(ShapeCount - 1).Height
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To MyDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Top = SlideShape(ShapeCount - 1).Top + SlideShape(ShapeCount - 1).Height
        Next ShapeCount
        
    End If
End Sub

Sub ObjectsIncreaseSpacingHorizontal()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As shape
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To MyDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).left = SlideShape(ShapeCount).left + (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    Else
        ReDim SlideShape(1 To MyDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).left = SlideShape(ShapeCount).left + (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    End If
    
End Sub

Sub ObjectsDecreaseSpacingHorizontal()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As shape
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To MyDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).left = SlideShape(ShapeCount).left - (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To MyDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByLeftPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).left = SlideShape(ShapeCount).left - (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    End If
End Sub

Sub ObjectsIncreaseSpacingVertical()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As shape
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To MyDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Top = SlideShape(ShapeCount).Top + (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To MyDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Top = SlideShape(ShapeCount).Top + (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    End If
    
End Sub

Sub ObjectsDecreaseSpacingVertical()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As shape
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        ReDim SlideShape(1 To MyDocument.Selection.ChildShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ChildShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ChildShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Top = SlideShape(ShapeCount).Top - (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    Else
        
        ReDim SlideShape(1 To MyDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To MyDocument.Selection.ShapeRange.Count
            Set SlideShape(ShapeCount) = MyDocument.Selection.ShapeRange(ShapeCount)
        Next ShapeCount
        
        ObjectsSortByTopPosition SlideShape
        
        For ShapeCount = 2 To UBound(SlideShape)
            SlideShape(ShapeCount).Top = SlideShape(ShapeCount).Top - (ShapeCount - 1) * 0.01 * 28.34646
        Next ShapeCount
        
    End If
    
End Sub

Sub ObjectsSortByLeftPosition(ArrayToSort As Variant)
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim StopLoop    As Boolean
    Dim ShapeCount  As Long
    Dim SlideShapes As shape
    Do
        StopLoop = False
        For ShapeCount = LBound(ArrayToSort) To UBound(ArrayToSort) - 1
            
            If ArrayToSort(ShapeCount).left > ArrayToSort(ShapeCount + 1).left Then
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
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim StopLoop    As Boolean
    Dim ShapeCount  As Long
    Dim SlideShapes As shape
    Do
        StopLoop = False
        For ShapeCount = LBound(ArrayToSort) To UBound(ArrayToSort) - 1
            
            If (ArrayToSort(ShapeCount).left + ArrayToSort(ShapeCount).Width) > (ArrayToSort(ShapeCount + 1).left + ArrayToSort(ShapeCount + 1).Width) Then
                Set SlideShapes = ArrayToSort(ShapeCount)
                Set ArrayToSort(ShapeCount) = ArrayToSort(ShapeCount + 1)
                Set ArrayToSort(ShapeCount + 1) = SlideShapes
                StopLoop = True
            End If
        Next ShapeCount
    Loop Until Not StopLoop
    
    Set SlideShapes = Nothing
End Sub


Sub ObjectsQuicksortTopLeftToBottomRight(SlideShapeRange As ShapeRange)
    
    If SlideShapeRange.Count = 0 Then Exit Sub
    
    Dim ArrayToSort() As shape
    ReDim ArrayToSort(1 To SlideShapeRange.Count)
    
    
    For i = 1 To SlideShapeRange.Count
        Set ArrayToSort(i) = SlideShapeRange(i)
    Next i
    
    QuicksortTopLeftToBottomRight ArrayToSort, 1, UBound(ArrayToSort)
    
    
    Dim NewSlideShapeRange As ShapeRange
    Dim ShapeNamesArray() As String
    ReDim ShapeNamesArray(1 To UBound(ArrayToSort))
    
    For i = 1 To UBound(ArrayToSort)
        ShapeNamesArray(i) = ArrayToSort(i).Name
    Next i
    
    Set NewSlideShapeRange = ActiveWindow.Selection.SlideRange.Shapes.Range(ShapeNamesArray)
    Set SlideShapeRange = NewSlideShapeRange

End Sub

Sub QuicksortTopLeftToBottomRight(ShapeItems() As shape, left As Long, right As Long)
    Dim i As Long, j As Long
    Dim PivotShape As shape
    Dim TempShape As shape
    
    i = left
    j = right
    
    Set PivotShape = ShapeItems((left + right) \ 2)
    Do
        Do While (ShapeItems(i).Top < PivotShape.Top) Or (ShapeItems(i).Top = PivotShape.Top And ShapeItems(i).left < PivotShape.left)
            i = i + 1
        Loop
        Do While (ShapeItems(j).Top > PivotShape.Top) Or (ShapeItems(j).Top = PivotShape.Top And ShapeItems(j).left > PivotShape.left)
            j = j - 1
        Loop
        If i <= j Then
            Set TempShape = ShapeItems(i)
            Set ShapeItems(i) = ShapeItems(j)
            Set ShapeItems(j) = TempShape
            i = i + 1
            j = j - 1
        End If
    Loop While i <= j
    
    If left < j Then QuicksortTopLeftToBottomRight ShapeItems, left, j
    
    If i < right Then QuicksortTopLeftToBottomRight ShapeItems, i, right
    
End Sub

Sub ObjectsSortByTopPosition(ArrayToSort As Variant)
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim StopLoop    As Boolean
    Dim ShapeCount  As Long
    Dim SlideShapes As shape
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
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim StopLoop    As Boolean
    Dim ShapeCount  As Long
    Dim SlideShapes As shape
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
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        If MyDocument.Selection.ChildShapeRange.Count > 1 Then
            MyDocument.Selection.ChildShapeRange.Align msoAlignLefts, msoFalse
        End If
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        MyDocument.Selection.ShapeRange.Align msoAlignLefts, msoTrue
    Else
        MyDocument.Selection.ShapeRange.Align msoAlignLefts, msoFalse
    End If
    
End Sub

Sub ObjectsAlignRights()
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        If MyDocument.Selection.ChildShapeRange.Count > 1 Then
            MyDocument.Selection.ChildShapeRange.Align msoAlignRights, msoFalse
        End If
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        MyDocument.Selection.ShapeRange.Align msoAlignRights, msoTrue
    Else
        MyDocument.Selection.ShapeRange.Align msoAlignRights, msoFalse
    End If
    
End Sub

Sub ObjectsAlignBottoms()
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        If MyDocument.Selection.ChildShapeRange.Count > 1 Then
            MyDocument.Selection.ChildShapeRange.Align msoAlignBottoms, msoFalse
        End If
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        MyDocument.Selection.ShapeRange.Align msoAlignBottoms, msoTrue
    Else
        MyDocument.Selection.ShapeRange.Align msoAlignBottoms, msoFalse
    End If
    
End Sub

Sub ObjectsAlignCenters()
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        If MyDocument.Selection.ChildShapeRange.Count > 1 Then
            MyDocument.Selection.ChildShapeRange.Align msoAlignCenters, msoFalse
        End If
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        MyDocument.Selection.ShapeRange.Align msoAlignCenters, msoTrue
    Else
        MyDocument.Selection.ShapeRange.Align msoAlignCenters, msoFalse
    End If
    
End Sub

Sub ObjectsAlignMiddles()
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        If MyDocument.Selection.ChildShapeRange.Count > 1 Then
            MyDocument.Selection.ChildShapeRange.Align msoAlignMiddles, msoFalse
        End If
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        MyDocument.Selection.ShapeRange.Align msoAlignMiddles, msoTrue
    Else
        MyDocument.Selection.ShapeRange.Align msoAlignMiddles, msoFalse
    End If
    
End Sub

Sub ObjectsAlignTops()
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        If MyDocument.Selection.ChildShapeRange.Count > 1 Then
            MyDocument.Selection.ChildShapeRange.Align msoAlignTops, msoFalse
        End If
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        MyDocument.Selection.ShapeRange.Align msoAlignTops, msoTrue
    Else
        MyDocument.Selection.ShapeRange.Align msoAlignTops, msoFalse
    End If
    
End Sub

Sub ObjectsDistributeHorizontally()
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.ShapeRange.Count > 2 Then
        MyDocument.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse
        
    ElseIf MyDocument.Selection.HasChildShapeRange Then
        
        If MyDocument.Selection.ChildShapeRange.Count > 2 Then
            MyDocument.Selection.ChildShapeRange.Distribute msoDistributeHorizontally, msoFalse
        End If
        
    Else
        MsgBox "Select more shapes to use this command."
    End If
    
End Sub

Sub ObjectsDistributeVertically()
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.ShapeRange.Count > 2 Then
        MyDocument.Selection.ShapeRange.Distribute msoDistributeVertically, msoFalse
        
    ElseIf MyDocument.Selection.HasChildShapeRange Then
        
        If MyDocument.Selection.ChildShapeRange.Count > 2 Then
            
            MyDocument.Selection.ChildShapeRange.Distribute msoDistributeVertically, msoFalse
            
        End If
        
    Else
        MsgBox "Select more shapes to use this command."
    End If
End Sub

Sub ArrangeShapes()
    Dim SlideShape         As shape
    Dim ShapeGroups      As Collection
    Set ShapeGroups = New Collection
    Dim SelectedShapeRange As ShapeRange
       
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Set SelectedShapeRange = MyDocument.Selection.ShapeRange
    
    NameRandomizer = Rnd(5)
     
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    For Each SlideShape In SelectedShapeRange
    SlideShape.Name = "Shape " & SlideShape.id
        
    Next SlideShape
    
    ObjectsQuicksortTopLeftToBottomRight SelectedShapeRange
       
    For Each SlideShape In SelectedShapeRange
        
        Dim ShapeShapeGroupExists As Boolean
        ShapeShapeGroupExists = False
        
        For Each ShapeGroup In ShapeGroups
            
            If (SlideShape.left + SlideShape.Width) >= ShapeGroup(1).left And SlideShape.left <= (ShapeGroup(1).left + ShapeGroup(1).Width) Then
                ShapeGroup.Add SlideShape
                ShapeShapeGroupExists = True
                Exit For
            End If
        Next ShapeGroup
        
        If Not ShapeShapeGroupExists Then
            Set NewShapeGroup = New Collection
            NewShapeGroup.Add SlideShape
            ShapeGroups.Add NewShapeGroup
        End If
    Next SlideShape
    
    Dim GroupNames() As String
    ReDim GroupNames(ShapeGroups.Count - 1)
    
    For i = 1 To ShapeGroups.Count
        GroupNames(i - 1) = "ColumnGroup" & i & " " & NameRandomizer
    Next i
    
    For Each ShapeGroup In ShapeGroups
        Dim ShapeNames() As String
        ReDim ShapeNames(ShapeGroup.Count - 1)
        
        a = a + 1
        
        For i = 1 To ShapeGroup.Count
            ShapeNames(i - 1) = ShapeGroup(i).Name
        Next i
        
        If ShapeGroup.Count > 1 Then
            MyDocument.Selection.SlideRange.Shapes.Range(ShapeNames).Align msoAlignCenters, msoFalse
            'MyDocument.Selection.SlideRange.Shapes.Range(ShapeNames).Distribute msoDistributeVertically, msoFalse
            Set ColumnGroup = MyDocument.Selection.SlideRange.Shapes.Range(ShapeNames).Group
            ColumnGroup.Name = "ColumnGroup" & a & " " & NameRandomizer
        Else
            
            Set ColumnGroup = MyDocument.Selection.SlideRange.Shapes.Range(ShapeNames)
            GroupNames(a - 1) = ShapeNames(0)
            ColumnGroup.Name = GroupNames(a - 1)
                        
        End If
    Next ShapeGroup
    
    If ShapeGroups.Count > 2 Then
    MyDocument.Selection.SlideRange.Shapes.Range(GroupNames).Distribute msoDistributeHorizontally, msoFalse
    End If
    
    MyDocument.Selection.SlideRange.Shapes.Range(GroupNames).Ungroup
    
    Set ShapeGroups = Nothing
    Set ShapeGroup = Nothing
    Set ShapeGroups = New Collection
    
    For Each SlideShape In SelectedShapeRange
        
        ShapeShapeGroupExists = False
        
        For Each ShapeGroup In ShapeGroups
            
            If (SlideShape.Top + SlideShape.Height) >= ShapeGroup(1).Top And SlideShape.Top <= (ShapeGroup(1).Top + ShapeGroup(1).Height) Then
                ShapeGroup.Add SlideShape
                ShapeShapeGroupExists = True
                Exit For
            End If
        Next ShapeGroup
        
        If Not ShapeShapeGroupExists Then
            Set NewShapeGroup = New Collection
            NewShapeGroup.Add SlideShape
            ShapeGroups.Add NewShapeGroup
        End If
    Next SlideShape
    
    a = 0
    Dim GroupNames2() As String
    ReDim GroupNames2(ShapeGroups.Count - 1)
    
    For i = 1 To ShapeGroups.Count
        GroupNames2(i - 1) = "RowGroup" & i & " " & NameRandomizer
    Next i
    
    For Each ShapeGroup In ShapeGroups
        Dim ShapeNames2() As String
        ReDim ShapeNames2(ShapeGroup.Count - 1)
        
        a = a + 1
        
        For i = 1 To ShapeGroup.Count
            ShapeNames2(i - 1) = ShapeGroup(i).Name
        Next i
        
        If ShapeGroup.Count > 1 Then
            MyDocument.Selection.SlideRange.Shapes.Range(ShapeNames2).Align msoAlignMiddles, msoFalse
            'MyDocument.Selection.SlideRange.Shapes.Range(ShapeNames2).Distribute msoDistributeHorizontally, msoFalse
            Set RowGroup = MyDocument.Selection.SlideRange.Shapes.Range(ShapeNames2).Group
            RowGroup.Name = "RowGroup" & a & " " & NameRandomizer
        Else
            
            Set RowGroup = MyDocument.Selection.SlideRange.Shapes.Range(ShapeNames2)
            GroupNames2(a - 1) = ShapeNames2(0)
            RowGroup.Name = GroupNames2(a - 1)
            
        End If
    Next ShapeGroup
    
    If ShapeGroups.Count > 2 Then
    MyDocument.Selection.SlideRange.Shapes.Range(GroupNames2).Distribute msoDistributeVertically, msoFalse
    End If
    
    MyDocument.Selection.SlideRange.Shapes.Range(GroupNames2).Ungroup
    
End Sub






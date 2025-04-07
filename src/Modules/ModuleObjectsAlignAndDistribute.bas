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
    Dim BaseTop As Single

    If MyDocument.Selection.HasChildShapeRange Then
        Set ShapeRange = MyDocument.Selection.ChildShapeRange
    Else
        Set ShapeRange = MyDocument.Selection.ShapeRange
    End If

    ReDim SlideShape(1 To ShapeRange.Count)

    For ShapeCount = 1 To ShapeRange.Count
        Set SlideShape(ShapeCount) = ShapeRange(ShapeCount)
    Next ShapeCount

    ObjectsSortByTopPosition SlideShape

    BaseTop = GetRealTop(SlideShape(1))

    For ShapeCount = 2 To UBound(SlideShape)
        SetRealHeight SlideShape(ShapeCount), GetRealHeight(SlideShape(ShapeCount)) + (GetRealTop(SlideShape(ShapeCount)) - BaseTop)
        SetRealTop SlideShape(ShapeCount), BaseTop
    Next ShapeCount
    
End Sub

Sub ObjectsStretchLeft()

    Set MyDocument = Application.ActiveWindow

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim ShapeCount As Long
    Dim SlideShape() As shape
    Dim BaseLeft As Single

    If MyDocument.Selection.HasChildShapeRange Then
        Set ShapeRange = MyDocument.Selection.ChildShapeRange
    Else
        Set ShapeRange = MyDocument.Selection.ShapeRange
    End If

    ReDim SlideShape(1 To ShapeRange.Count)

    For ShapeCount = 1 To ShapeRange.Count
        Set SlideShape(ShapeCount) = ShapeRange(ShapeCount)
    Next ShapeCount

    ObjectsSortByLeftPosition SlideShape

    BaseLeft = GetRealLeft(SlideShape(1))

    For ShapeCount = 2 To UBound(SlideShape)
        SetRealWidth SlideShape(ShapeCount), GetRealWidth(SlideShape(ShapeCount)) + (GetRealLeft(SlideShape(ShapeCount)) - BaseLeft)
        SetRealLeft SlideShape(ShapeCount), BaseLeft
    Next ShapeCount
End Sub

Sub ObjectsStretchBottom()

    Set MyDocument = Application.ActiveWindow

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim ShapeCount As Long
    Dim SlideShape() As shape
    Dim BaseBottom As Single

    If MyDocument.Selection.HasChildShapeRange Then
        Set ShapeRange = MyDocument.Selection.ChildShapeRange
    Else
        Set ShapeRange = MyDocument.Selection.ShapeRange
    End If

    ReDim SlideShape(1 To ShapeRange.Count)

    For ShapeCount = 1 To ShapeRange.Count
        Set SlideShape(ShapeCount) = ShapeRange(ShapeCount)
    Next ShapeCount

    ObjectsSortByBottomPosition SlideShape

    BaseBottom = GetRealTop(SlideShape(UBound(SlideShape))) + GetRealHeight(SlideShape(UBound(SlideShape)))

    For ShapeCount = UBound(SlideShape) - 1 To 1 Step -1
        SetRealHeight SlideShape(ShapeCount), GetRealHeight(SlideShape(ShapeCount)) + (BaseBottom - GetRealTop(SlideShape(ShapeCount)) - GetRealHeight(SlideShape(ShapeCount)))
    Next ShapeCount
End Sub

Sub ObjectsStretchRight()

    Set MyDocument = Application.ActiveWindow

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim ShapeCount As Long
    Dim SlideShape() As shape
    Dim BaseRight As Single

    If MyDocument.Selection.HasChildShapeRange Then
        Set ShapeRange = MyDocument.Selection.ChildShapeRange
    Else
        Set ShapeRange = MyDocument.Selection.ShapeRange
    End If

    ReDim SlideShape(1 To ShapeRange.Count)

    For ShapeCount = 1 To ShapeRange.Count
        Set SlideShape(ShapeCount) = ShapeRange(ShapeCount)
    Next ShapeCount

    ObjectsSortByRightPosition SlideShape

    BaseRight = GetRealLeft(SlideShape(UBound(SlideShape))) + GetRealWidth(SlideShape(UBound(SlideShape)))

    For ShapeCount = UBound(SlideShape) - 1 To 1 Step -1
        SetRealWidth SlideShape(ShapeCount), GetRealWidth(SlideShape(ShapeCount)) + (BaseRight - GetRealLeft(SlideShape(ShapeCount)) - GetRealWidth(SlideShape(ShapeCount)))
    Next ShapeCount
End Sub

Sub ObjectsStretchTopShapeBottom()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As shape
    Dim BaseTop As Single

    If MyDocument.Selection.HasChildShapeRange Then
        Set ShapeRange = MyDocument.Selection.ChildShapeRange
    Else
        Set ShapeRange = MyDocument.Selection.ShapeRange
    End If

    ReDim SlideShape(1 To ShapeRange.Count)

    For ShapeCount = 1 To ShapeRange.Count
        Set SlideShape(ShapeCount) = ShapeRange(ShapeCount)
    Next ShapeCount

    ObjectsSortByBottomPosition SlideShape

    BaseTop = GetRealTop(SlideShape(1)) + GetRealHeight(SlideShape(1))

    For ShapeCount = 2 To UBound(SlideShape)
        SetRealHeight SlideShape(ShapeCount), GetRealHeight(SlideShape(ShapeCount)) + (GetRealTop(SlideShape(ShapeCount)) - BaseTop)
        SetRealTop SlideShape(ShapeCount), BaseTop
    Next ShapeCount
    
End Sub

Sub ObjectsStretchLeftShapeRight()

    Set MyDocument = Application.ActiveWindow

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim ShapeCount As Long
    Dim SlideShape() As shape
    Dim BaseLeft As Single

    If MyDocument.Selection.HasChildShapeRange Then
        Set ShapeRange = MyDocument.Selection.ChildShapeRange
    Else
        Set ShapeRange = MyDocument.Selection.ShapeRange
    End If

    ReDim SlideShape(1 To ShapeRange.Count)

    For ShapeCount = 1 To ShapeRange.Count
        Set SlideShape(ShapeCount) = ShapeRange(ShapeCount)
    Next ShapeCount

    ObjectsSortByRightPosition SlideShape

    BaseLeft = GetRealLeft(SlideShape(1)) + GetRealWidth(SlideShape(1))

    For ShapeCount = 2 To UBound(SlideShape)
        SetRealWidth SlideShape(ShapeCount), GetRealWidth(SlideShape(ShapeCount)) + (GetRealLeft(SlideShape(ShapeCount)) - BaseLeft)
        SetRealLeft SlideShape(ShapeCount), BaseLeft
    Next ShapeCount
End Sub

Sub ObjectsStretchBottomShapeTop()

    Set MyDocument = Application.ActiveWindow

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim ShapeCount As Long
    Dim SlideShape() As shape
    Dim BaseBottom As Single

    If MyDocument.Selection.HasChildShapeRange Then
        Set ShapeRange = MyDocument.Selection.ChildShapeRange
    Else
        Set ShapeRange = MyDocument.Selection.ShapeRange
    End If

    ReDim SlideShape(1 To ShapeRange.Count)

    For ShapeCount = 1 To ShapeRange.Count
        Set SlideShape(ShapeCount) = ShapeRange(ShapeCount)
    Next ShapeCount

    ObjectsSortByTopPosition SlideShape

    BaseBottom = GetRealTop(SlideShape(UBound(SlideShape)))

    For ShapeCount = UBound(SlideShape) - 1 To 1 Step -1
        SetRealHeight SlideShape(ShapeCount), GetRealHeight(SlideShape(ShapeCount)) + (BaseBottom - GetRealTop(SlideShape(ShapeCount)) - GetRealHeight(SlideShape(ShapeCount)))
    Next ShapeCount
End Sub

Sub ObjectsStretchRightShapeLeft()

    Set MyDocument = Application.ActiveWindow

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim ShapeCount As Long
    Dim SlideShape() As shape
    Dim BaseRight As Single

    If MyDocument.Selection.HasChildShapeRange Then
        Set ShapeRange = MyDocument.Selection.ChildShapeRange
    Else
        Set ShapeRange = MyDocument.Selection.ShapeRange
    End If

    ReDim SlideShape(1 To ShapeRange.Count)

    For ShapeCount = 1 To ShapeRange.Count
        Set SlideShape(ShapeCount) = ShapeRange(ShapeCount)
    Next ShapeCount

    ObjectsSortByLeftPosition SlideShape

    BaseRight = GetRealLeft(SlideShape(UBound(SlideShape)))

    For ShapeCount = UBound(SlideShape) - 1 To 1 Step -1
        SetRealWidth SlideShape(ShapeCount), GetRealWidth(SlideShape(ShapeCount)) + (BaseRight - GetRealLeft(SlideShape(ShapeCount)) - GetRealWidth(SlideShape(ShapeCount)))
    Next ShapeCount
End Sub

Sub ObjectsRemoveSpacingHorizontal()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As shape

    If MyDocument.Selection.HasChildShapeRange Then
        Set ShapeRange = MyDocument.Selection.ChildShapeRange
    Else
        Set ShapeRange = MyDocument.Selection.ShapeRange
    End If

    ReDim SlideShape(1 To ShapeRange.Count)

    For ShapeCount = 1 To ShapeRange.Count
        Set SlideShape(ShapeCount) = ShapeRange(ShapeCount)
    Next ShapeCount

    ObjectsSortByLeftPosition SlideShape

    For ShapeCount = 2 To UBound(SlideShape)
        SetRealLeft SlideShape(ShapeCount), GetRealLeft(SlideShape(ShapeCount - 1)) + GetRealWidth(SlideShape(ShapeCount - 1))
    Next ShapeCount
End Sub

Sub ObjectsRemoveSpacingVertical()

    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim ShapeCount  As Long
    Dim SlideShape() As shape

    If MyDocument.Selection.HasChildShapeRange Then
        Set ShapeRange = MyDocument.Selection.ChildShapeRange
    Else
        Set ShapeRange = MyDocument.Selection.ShapeRange
    End If

    ReDim SlideShape(1 To ShapeRange.Count)

    For ShapeCount = 1 To ShapeRange.Count
        Set SlideShape(ShapeCount) = ShapeRange(ShapeCount)
    Next ShapeCount

    ObjectsSortByTopPosition SlideShape

    For ShapeCount = 2 To UBound(SlideShape)
        SetRealTop SlideShape(ShapeCount), GetRealTop(SlideShape(ShapeCount - 1)) + GetRealHeight(SlideShape(ShapeCount - 1))
    Next ShapeCount
End Sub

Sub ObjectsIncreaseSpacingHorizontal()

    Set MyDocument = Application.ActiveWindow

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim ShapeCount As Long
    Dim SlideShape() As shape

    If MyDocument.Selection.HasChildShapeRange Then
        Set ShapeRange = MyDocument.Selection.ChildShapeRange
    Else
        Set ShapeRange = MyDocument.Selection.ShapeRange
    End If

    ReDim SlideShape(1 To ShapeRange.Count)

    For ShapeCount = 1 To ShapeRange.Count
        Set SlideShape(ShapeCount) = ShapeRange(ShapeCount)
    Next ShapeCount

    ObjectsSortByLeftPosition SlideShape

    For ShapeCount = 2 To UBound(SlideShape)
        SetRealLeft SlideShape(ShapeCount), GetRealLeft(SlideShape(ShapeCount)) + (ShapeCount - 1) * 0.01 * 28.34646
    Next ShapeCount
End Sub

Sub ObjectsDecreaseSpacingHorizontal()

    Set MyDocument = Application.ActiveWindow

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim ShapeCount As Long
    Dim SlideShape() As shape

    If MyDocument.Selection.HasChildShapeRange Then
        Set ShapeRange = MyDocument.Selection.ChildShapeRange
    Else
        Set ShapeRange = MyDocument.Selection.ShapeRange
    End If

    ReDim SlideShape(1 To ShapeRange.Count)

    For ShapeCount = 1 To ShapeRange.Count
        Set SlideShape(ShapeCount) = ShapeRange(ShapeCount)
    Next ShapeCount

    ObjectsSortByLeftPosition SlideShape

    For ShapeCount = 2 To UBound(SlideShape)
        SetRealLeft SlideShape(ShapeCount), GetRealLeft(SlideShape(ShapeCount)) - (ShapeCount - 1) * 0.01 * 28.34646
    Next ShapeCount
End Sub

Sub ObjectsIncreaseSpacingVertical()

    Set MyDocument = Application.ActiveWindow

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim ShapeCount As Long
    Dim SlideShape() As shape

    If MyDocument.Selection.HasChildShapeRange Then
        Set ShapeRange = MyDocument.Selection.ChildShapeRange
    Else
        Set ShapeRange = MyDocument.Selection.ShapeRange
    End If

    ReDim SlideShape(1 To ShapeRange.Count)

    For ShapeCount = 1 To ShapeRange.Count
        Set SlideShape(ShapeCount) = ShapeRange(ShapeCount)
    Next ShapeCount

    ObjectsSortByTopPosition SlideShape

    For ShapeCount = 2 To UBound(SlideShape)
        SetRealTop SlideShape(ShapeCount), GetRealTop(SlideShape(ShapeCount)) + (ShapeCount - 1) * 0.01 * 28.34646
    Next ShapeCount
End Sub

Sub ObjectsDecreaseSpacingVertical()

    Set MyDocument = Application.ActiveWindow

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim ShapeCount As Long
    Dim SlideShape() As shape

    If MyDocument.Selection.HasChildShapeRange Then
        Set ShapeRange = MyDocument.Selection.ChildShapeRange
    Else
        Set ShapeRange = MyDocument.Selection.ShapeRange
    End If

    ReDim SlideShape(1 To ShapeRange.Count)

    For ShapeCount = 1 To ShapeRange.Count
        Set SlideShape(ShapeCount) = ShapeRange(ShapeCount)
    Next ShapeCount

    ObjectsSortByTopPosition SlideShape

    For ShapeCount = 2 To UBound(SlideShape)
        SetRealTop SlideShape(ShapeCount), GetRealTop(SlideShape(ShapeCount)) - (ShapeCount - 1) * 0.01 * 28.34646
    Next ShapeCount
End Sub

Sub ObjectsSortByLeftPosition(ArrayToSort() As shape)
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim StopLoop    As Boolean
    Dim ShapeCount  As Long
    Dim SlideShapes As shape
    Do
        StopLoop = False
        For ShapeCount = LBound(ArrayToSort) To UBound(ArrayToSort) - 1
            
            If GetRealLeft(ArrayToSort(ShapeCount)) > GetRealLeft(ArrayToSort(ShapeCount + 1)) Then
                Set SlideShapes = ArrayToSort(ShapeCount)
                Set ArrayToSort(ShapeCount) = ArrayToSort(ShapeCount + 1)
                Set ArrayToSort(ShapeCount + 1) = SlideShapes
                StopLoop = True
            End If
        Next ShapeCount
    Loop Until Not StopLoop
    
    Set SlideShapes = Nothing
End Sub

Sub ObjectsSortByRightPosition(ArrayToSort() As shape)
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim StopLoop    As Boolean
    Dim ShapeCount  As Long
    Dim SlideShapes As shape
    Do
        StopLoop = False
        For ShapeCount = LBound(ArrayToSort) To UBound(ArrayToSort) - 1
            
            If (GetRealLeft(ArrayToSort(ShapeCount)) + GetRealWidth(ArrayToSort(ShapeCount))) > (GetRealLeft(ArrayToSort(ShapeCount + 1)) + GetRealWidth(ArrayToSort(ShapeCount + 1))) Then
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
        Do While (GetRealTop(ShapeItems(i)) < GetRealTop(PivotShape)) Or (GetRealTop(ShapeItems(i)) = GetRealTop(PivotShape) And GetRealLeft(ShapeItems(i)) < GetRealLeft(PivotShape))
            i = i + 1
        Loop
        Do While (GetRealTop(ShapeItems(j)) > GetRealTop(PivotShape)) Or (GetRealTop(ShapeItems(j)) = GetRealTop(PivotShape) And GetRealLeft(ShapeItems(j)) > GetRealLeft(PivotShape))
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

Sub ObjectsSortByTopPosition(ArrayToSort() As shape)
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim StopLoop    As Boolean
    Dim ShapeCount  As Long
    Dim SlideShapes As shape
    Do
        StopLoop = False
        For ShapeCount = LBound(ArrayToSort) To UBound(ArrayToSort) - 1
            
            If GetRealTop(ArrayToSort(ShapeCount)) > GetRealTop(ArrayToSort(ShapeCount + 1)) Then
                Set SlideShapes = ArrayToSort(ShapeCount)
                Set ArrayToSort(ShapeCount) = ArrayToSort(ShapeCount + 1)
                Set ArrayToSort(ShapeCount + 1) = SlideShapes
                StopLoop = True
            End If
        Next ShapeCount
    Loop Until Not StopLoop
    
    Set SlideShapes = Nothing
End Sub

Sub ObjectsSortByBottomPosition(ArrayToSort() As shape)
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim StopLoop    As Boolean
    Dim ShapeCount  As Long
    Dim SlideShapes As shape
    Do
        StopLoop = False
        For ShapeCount = LBound(ArrayToSort) To UBound(ArrayToSort) - 1
            
            If (GetRealTop(ArrayToSort(ShapeCount)) + GetRealHeight(ArrayToSort(ShapeCount))) > (GetRealTop(ArrayToSort(ShapeCount + 1)) + GetRealHeight(ArrayToSort(ShapeCount + 1))) Then
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
    Dim Left1 As Single
    Dim Left2 As Single
    Dim sh As shape
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        If MyDocument.Selection.ChildShapeRange.Count > 1 Then
            
            Left1 = GetRealLeft(MyDocument.Selection.ChildShapeRange(1))
            Left2 = GetRealLeft(MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.Count))
            MyDocument.Selection.ChildShapeRange.Align msoAlignLefts, msoFalse
        End If
        
        If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 1 Then
            For Each sh In MyDocument.Selection.ChildShapeRange
                SetRealLeft sh, Left1
            Next sh
        End If
        
        If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 2 Then
            For Each sh In MyDocument.Selection.ChildShapeRange
                SetRealLeft sh, Left2
            Next sh
        End If
        
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        MyDocument.Selection.ShapeRange.Align msoAlignLefts, msoTrue
    Else
        
        Left1 = GetRealLeft(MyDocument.Selection.ShapeRange(1))
        Left2 = GetRealLeft(MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.Count))
        
        MyDocument.Selection.ShapeRange.Align msoAlignLefts, msoFalse
        
        If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 1 Then
            For Each sh In MyDocument.Selection.ShapeRange
                SetRealLeft sh, Left1
            Next sh
        End If
        
        If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 2 Then
            For Each sh In MyDocument.Selection.ShapeRange
                SetRealLeft sh, Left2
            Next sh
        End If
        
    End If
    
End Sub

Sub ObjectsAlignTops()
    Set MyDocument = Application.ActiveWindow
    Dim Top1 As Single
    Dim Top2 As Single
    Dim sh As shape
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        If MyDocument.Selection.ChildShapeRange.Count > 1 Then
            
            Top1 = GetRealTop(MyDocument.Selection.ChildShapeRange(1))
            Top2 = GetRealTop(MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.Count))
            
            MyDocument.Selection.ChildShapeRange.Align msoAlignTops, msoFalse
        End If
        
        If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 1 Then
            For Each sh In MyDocument.Selection.ChildShapeRange
                SetRealTop sh, Top1
            Next sh
        End If
        
        If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 2 Then
            For Each sh In MyDocument.Selection.ChildShapeRange
                SetRealTop sh, Top2
            Next sh
        End If
        
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        MyDocument.Selection.ShapeRange.Align msoAlignTops, msoTrue
    Else
        
        Top1 = GetRealTop(MyDocument.Selection.ShapeRange(1))
        Top2 = GetRealTop(MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.Count))
        
        If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 0 Then
            
        MyDocument.Selection.ShapeRange.Align msoAlignTops, msoFalse
        
        ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 1 Then
            For Each sh In MyDocument.Selection.ShapeRange
                SetRealTop sh, Top1
            Next sh
        End If
        
        If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 2 Then
            For Each sh In MyDocument.Selection.ShapeRange
                SetRealTop sh, Top2
            Next sh
        End If
        
    End If
    
End Sub

Sub ObjectsAlignRights()
    Set MyDocument = Application.ActiveWindow
    
    Dim Right1 As Single
    Dim Right2 As Single
    Dim sh As shape
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        If MyDocument.Selection.ChildShapeRange.Count > 1 Then
            
            Right1 = GetRealLeft(MyDocument.Selection.ChildShapeRange(1)) + GetRealWidth(MyDocument.Selection.ChildShapeRange(1))
            Right2 = GetRealLeft(MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.Count)) + GetRealWidth(MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.Count))
            
            If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 0 Then
                
                MyDocument.Selection.ChildShapeRange.Align msoAlignRights, msoFalse
                
            ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 1 Then
                
                For Each sh In MyDocument.Selection.ChildShapeRange
                    SetRealLeft sh, Right1 - GetRealWidth(sh)
                Next sh
                
            ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 2 Then
                
                For Each sh In MyDocument.Selection.ChildShapeRange
                    SetRealLeft sh, Right2 - GetRealWidth(sh)
                Next sh
                
            End If
            
        End If
        
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        MyDocument.Selection.ShapeRange.Align msoAlignRights, msoTrue
    Else
        
        Right1 = GetRealLeft(MyDocument.Selection.ShapeRange(1)) + GetRealWidth(MyDocument.Selection.ShapeRange(1))
        Right2 = GetRealLeft(MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.Count)) + GetRealWidth(MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.Count))
        
        If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 0 Then
            
            MyDocument.Selection.ShapeRange.Align msoAlignRights, msoFalse
            
        ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 1 Then
            
            For Each sh In MyDocument.Selection.ShapeRange
                SetRealLeft sh, Right1 - GetRealWidth(sh)
            Next sh
            
        ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 2 Then
            
            For Each sh In MyDocument.Selection.ShapeRange
                SetRealLeft sh, Right2 - GetRealWidth(sh)
            Next sh
            
        End If
        
    End If
    
End Sub

Sub ObjectsAlignBottoms()
    Set MyDocument = Application.ActiveWindow
    
    Dim Bottom1 As Single
    Dim Bottom2 As Single
    Dim sh As shape
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        If MyDocument.Selection.ChildShapeRange.Count > 1 Then
            
            Bottom1 = GetRealTop(MyDocument.Selection.ChildShapeRange(1)) + GetRealHeight(MyDocument.Selection.ChildShapeRange(1))
            Bottom2 = GetRealTop(MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.Count)) + GetRealHeight(MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.Count))
            
            If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 0 Then
                
                MyDocument.Selection.ChildShapeRange.Align msoAlignBottoms, msoFalse
                
            ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 1 Then
                
                For Each sh In MyDocument.Selection.ChildShapeRange
                    SetRealTop sh, Bottom1 - GetRealHeight(sh)
                Next sh
                
            ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 2 Then
                
                For Each sh In MyDocument.Selection.ChildShapeRange
                    SetRealTop sh, Bottom2 - GetRealHeight(sh)
                Next sh
                
            End If
            
        End If
        
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        MyDocument.Selection.ShapeRange.Align msoAlignBottoms, msoTrue
    Else
        
        Bottom1 = GetRealTop(MyDocument.Selection.ShapeRange(1)) + GetRealHeight(MyDocument.Selection.ShapeRange(1))
        Bottom2 = GetRealTop(MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.Count)) + GetRealHeight(MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.Count))
        
        If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 0 Then
            
            MyDocument.Selection.ShapeRange.Align msoAlignBottoms, msoFalse
            
        ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 1 Then
            
            For Each sh In MyDocument.Selection.ShapeRange
                SetRealTop sh, Bottom1 - GetRealHeight(sh)
            Next sh
            
        ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 2 Then
            
            For Each sh In MyDocument.Selection.ShapeRange
                SetRealTop sh, Bottom2 - GetRealHeight(sh)
            Next sh
            
        End If
        
    End If
    
End Sub

Sub ObjectsAlignCenters()
    Set MyDocument = Application.ActiveWindow
    Dim Center1, Center2 As Single
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        If MyDocument.Selection.ChildShapeRange.Count > 1 Then
            
            Center1 = MyDocument.Selection.ChildShapeRange(1).left + (MyDocument.Selection.ChildShapeRange(1).Width / 2)
            Center2 = MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.Count).left + (MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.Count).Width / 2)
            
            If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 0 Then
                
                MyDocument.Selection.ChildShapeRange.Align msoAlignCenters, msoFalse
                
            ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 1 Then
                
                For Each SlideShape In MyDocument.Selection.ChildShapeRange
                    SlideShape.left = Center1 - (SlideShape.Width / 2)
                Next SlideShape
                
            ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 2 Then
                
                For Each SlideShape In MyDocument.Selection.ChildShapeRange
                    SlideShape.left = Center2 - (SlideShape.Width / 2)
                Next SlideShape
                
            End If
            
        End If
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        MyDocument.Selection.ShapeRange.Align msoAlignCenters, msoTrue
    Else
        
        Center1 = MyDocument.Selection.ShapeRange(1).left + (MyDocument.Selection.ShapeRange(1).Width / 2)
        Center2 = MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.Count).left + (MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.Count).Width / 2)
        
        If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 0 Then
            
            MyDocument.Selection.ShapeRange.Align msoAlignCenters, msoFalse
            
        ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 1 Then
            
            For Each SlideShape In MyDocument.Selection.ShapeRange
                SlideShape.left = Center1 - (SlideShape.Width / 2)
            Next SlideShape
            
        ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 2 Then
            
            For Each SlideShape In MyDocument.Selection.ShapeRange
                SlideShape.left = Center2 - (SlideShape.Width / 2)
            Next SlideShape
            
        End If
        
    End If
    
End Sub

Sub ObjectsAlignMiddles()
    Set MyDocument = Application.ActiveWindow
    Dim Middle1, Middle2 As Single
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        If MyDocument.Selection.ChildShapeRange.Count > 1 Then
            
            Middle1 = MyDocument.Selection.ChildShapeRange(1).Top + (MyDocument.Selection.ChildShapeRange(1).Height / 2)
            Middle2 = MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.Count).Top + (MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.Count).Height / 2)
            
            If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 0 Then
                
                MyDocument.Selection.ChildShapeRange.Align msoAlignMiddles, msoFalse
                
            ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 1 Then
                
                For Each SlideShape In MyDocument.Selection.ChildShapeRange
                    SlideShape.Top = Middle1 - (SlideShape.Height / 2)
                Next SlideShape
                
            ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 2 Then
                
                For Each SlideShape In MyDocument.Selection.ChildShapeRange
                    SlideShape.Top = Middle2 - (SlideShape.Height / 2)
                Next SlideShape
                
            End If
            
        End If
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        MyDocument.Selection.ShapeRange.Align msoAlignMiddles, msoTrue
    Else
        
        Middle1 = MyDocument.Selection.ShapeRange(1).Top + (MyDocument.Selection.ShapeRange(1).Height / 2)
        Middle2 = MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.Count).Top + (MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.Count).Height / 2)
        
        If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 0 Then
            
            MyDocument.Selection.ShapeRange.Align msoAlignMiddles, msoFalse
            
        ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 1 Then
            
            For Each SlideShape In MyDocument.Selection.ShapeRange
                SlideShape.Top = Middle1 - (SlideShape.Height / 2)
            Next SlideShape
            
        ElseIf GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0") = 2 Then
            
            For Each SlideShape In MyDocument.Selection.ShapeRange
                SlideShape.Top = Middle2 - (SlideShape.Height / 2)
            Next SlideShape
            
        End If
        
    End If
    
End Sub


Sub ObjectsDistributeHorizontally()
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.ShapeRange.Count > 2 Then
        MyDocument.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
     MyDocument.Selection.ShapeRange.Align msoAlignCenters, msoTrue
        
    ElseIf MyDocument.Selection.HasChildShapeRange Then
        
        If MyDocument.Selection.ChildShapeRange.Count > 2 Then
            MyDocument.Selection.ChildShapeRange.Distribute msoDistributeHorizontally, msoFalse
        End If
    End If
    
End Sub

Sub ObjectsDistributeVertically()
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.ShapeRange.Count > 2 Then
        MyDocument.Selection.ShapeRange.Distribute msoDistributeVertically, msoFalse
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
    
     MyDocument.Selection.ShapeRange.Align msoAlignMiddles, msoTrue
    
    ElseIf MyDocument.Selection.HasChildShapeRange Then
        
        If MyDocument.Selection.ChildShapeRange.Count > 2 Then
            
            MyDocument.Selection.ChildShapeRange.Distribute msoDistributeVertically, msoFalse
            
        End If
    End If
End Sub

Sub ArrangeShapes()
    Dim SlideShape         As shape
    Dim ShapeGroups      As Collection
    Dim FirstShape As shape
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
            
            Set FirstShape = ShapeGroup(1)

            If (GetRealLeft(SlideShape) + GetRealWidth(SlideShape)) >= GetRealLeft(FirstShape) _
            And GetRealLeft(SlideShape) <= (GetRealLeft(FirstShape) + GetRealWidth(FirstShape)) Then
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
            
            Set FirstShape = ShapeGroup(1)

            If (GetRealTop(SlideShape) + GetRealHeight(SlideShape)) >= GetRealTop(FirstShape) _
            And GetRealTop(SlideShape) <= (GetRealTop(FirstShape) + GetRealHeight(FirstShape)) Then
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

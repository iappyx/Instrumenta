Attribute VB_Name = "ModuleObjectsSizeAndPosition"
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

Sub ObjectsSizeToTallest()
    Set MyDocument = Application.ActiveWindow
    Dim Height As Single
    Dim Tallest As Single
    Dim SlideShape As Shape

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        Tallest = GetRealHeight(MyDocument.Selection.ChildShapeRange(1))
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            Height = GetRealHeight(SlideShape)
            If Height > Tallest Then Tallest = Height
        Next
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            SetRealHeight SlideShape, Tallest
        Next SlideShape
        
    Else
        Tallest = GetRealHeight(MyDocument.Selection.shapeRange(1))
        
        For Each SlideShape In MyDocument.Selection.shapeRange
            Height = GetRealHeight(SlideShape)
            If Height > Tallest Then Tallest = Height
        Next
        
        For Each SlideShape In MyDocument.Selection.shapeRange
            SetRealHeight SlideShape, Tallest
        Next SlideShape
        
    End If
    
End Sub

Sub ObjectsSizeToShortest()
    Set MyDocument = Application.ActiveWindow
    Dim Height As Single
    Dim Shortest As Single
    Dim SlideShape As Shape

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    If MyDocument.Selection.HasChildShapeRange Then

        Shortest = GetRealHeight(MyDocument.Selection.ChildShapeRange(1))

        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            Height = GetRealHeight(SlideShape)
            If Height < Shortest Then Shortest = Height
        Next

        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            SetRealHeight SlideShape, Shortest
        Next SlideShape


    Else

        Shortest = GetRealHeight(MyDocument.Selection.shapeRange(1))

        For Each SlideShape In MyDocument.Selection.shapeRange
            Height = GetRealHeight(SlideShape)
            If Height < Shortest Then Shortest = Height
        Next

        For Each SlideShape In MyDocument.Selection.shapeRange
            SetRealHeight SlideShape, Shortest
        Next SlideShape

    End If

End Sub

Sub ObjectsSizeToWidest()
    Set MyDocument = Application.ActiveWindow
    Dim Width As Single
    Dim Widest As Single
    Dim SlideShape As Shape

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    If MyDocument.Selection.HasChildShapeRange Then

        Widest = GetRealWidth(MyDocument.Selection.ChildShapeRange(1))

        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            Width = GetRealWidth(SlideShape)
            If Width > Widest Then Widest = Width
        Next SlideShape

        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            SetRealWidth SlideShape, Widest
        Next SlideShape

    Else

        Widest = GetRealWidth(MyDocument.Selection.shapeRange(1))

        For Each SlideShape In MyDocument.Selection.shapeRange
            Width = GetRealWidth(SlideShape)
            If Width > Widest Then Widest = Width
        Next SlideShape

        For Each SlideShape In MyDocument.Selection.shapeRange
            SetRealWidth SlideShape, Widest
        Next SlideShape

    End If

End Sub

Sub ObjectsSizeToNarrowest()
    Set MyDocument = Application.ActiveWindow
    Dim Width As Single
    Dim Narrowest As Single
    Dim SlideShape As Shape

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    If MyDocument.Selection.HasChildShapeRange Then

        Narrowest = GetRealWidth(MyDocument.Selection.ChildShapeRange(1))

        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            Width = GetRealWidth(SlideShape)
            If Width < Narrowest Then Narrowest = Width
        Next SlideShape

        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            SetRealWidth SlideShape, Narrowest
        Next SlideShape

    Else

        Narrowest = GetRealWidth(MyDocument.Selection.shapeRange(1))

        For Each SlideShape In MyDocument.Selection.shapeRange
            Width = GetRealWidth(SlideShape)
            If Width < Narrowest Then Narrowest = Width
        Next SlideShape

        For Each SlideShape In MyDocument.Selection.shapeRange
            SetRealWidth SlideShape, Narrowest
        Next SlideShape

    End If

End Sub

Sub ObjectsSameHeight()
    Set MyDocument = Application.ActiveWindow
    Dim Height As Single
    Dim SlideShape As Shape

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultTransformationMethod", "0") = 0 Then
        If MyDocument.Selection.HasChildShapeRange Then
            Height = GetRealHeight(MyDocument.Selection.ChildShapeRange(1))

            For Each SlideShape In MyDocument.Selection.ChildShapeRange
                SetRealHeight SlideShape, Height
            Next SlideShape
        Else
            Height = GetRealHeight(MyDocument.Selection.shapeRange(1))

            For Each SlideShape In MyDocument.Selection.shapeRange
                SetRealHeight SlideShape, Height
            Next SlideShape
        End If
    Else
        If MyDocument.Selection.HasChildShapeRange Then
            Height = GetRealHeight(MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.Count))

            For Each SlideShape In MyDocument.Selection.ChildShapeRange
                SetRealHeight SlideShape, Height
            Next SlideShape
        Else
            Height = GetRealHeight(MyDocument.Selection.shapeRange(MyDocument.Selection.shapeRange.Count))

            For Each SlideShape In MyDocument.Selection.shapeRange
                SetRealHeight SlideShape, Height
            Next SlideShape
        End If
    End If
End Sub

Sub ObjectsSameWidth()
    Set MyDocument = Application.ActiveWindow
    Dim Width As Single
    Dim SlideShape As Shape

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultTransformationMethod", "0") = 0 Then
        If MyDocument.Selection.HasChildShapeRange Then
            Width = GetRealWidth(MyDocument.Selection.ChildShapeRange(1))

            For Each SlideShape In MyDocument.Selection.ChildShapeRange
                SetRealWidth SlideShape, Width
            Next SlideShape
        Else
            Width = GetRealWidth(MyDocument.Selection.shapeRange(1))

            For Each SlideShape In MyDocument.Selection.shapeRange
                SetRealWidth SlideShape, Width
            Next SlideShape
        End If
    Else
        If MyDocument.Selection.HasChildShapeRange Then
            Width = GetRealWidth(MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.Count))

            For Each SlideShape In MyDocument.Selection.ChildShapeRange
                SetRealWidth SlideShape, Width
            Next SlideShape
        Else
            Width = GetRealWidth(MyDocument.Selection.shapeRange(MyDocument.Selection.shapeRange.Count))

            For Each SlideShape In MyDocument.Selection.shapeRange
                SetRealWidth SlideShape, Width
            Next SlideShape
        End If
    End If
End Sub

Sub ObjectsSameHeightAndWidth()
    Set MyDocument = Application.ActiveWindow
    Dim Height As Single
    Dim Width As Single
    Dim SlideShape As Shape

    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub

    If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultTransformationMethod", "0") = 0 Then
        If MyDocument.Selection.HasChildShapeRange Then
            Height = GetRealHeight(MyDocument.Selection.ChildShapeRange(1))
            Width = GetRealWidth(MyDocument.Selection.ChildShapeRange(1))

            For Each SlideShape In MyDocument.Selection.ChildShapeRange
                SetRealHeight SlideShape, Height
                SetRealWidth SlideShape, Width
            Next SlideShape
        Else
            Height = GetRealHeight(MyDocument.Selection.shapeRange(1))
            Width = GetRealWidth(MyDocument.Selection.shapeRange(1))

            For Each SlideShape In MyDocument.Selection.shapeRange
                SetRealHeight SlideShape, Height
                SetRealWidth SlideShape, Width
            Next SlideShape
        End If
    Else
        If MyDocument.Selection.HasChildShapeRange Then
            Height = GetRealHeight(MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.Count))
            Width = GetRealWidth(MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.Count))

            For Each SlideShape In MyDocument.Selection.ChildShapeRange
                SetRealHeight SlideShape, Height
                SetRealWidth SlideShape, Width
            Next SlideShape
        Else
            Height = GetRealHeight(MyDocument.Selection.shapeRange(MyDocument.Selection.shapeRange.Count))
            Width = GetRealWidth(MyDocument.Selection.shapeRange(MyDocument.Selection.shapeRange.Count))

            For Each SlideShape In MyDocument.Selection.shapeRange
                SetRealHeight SlideShape, Height
                SetRealWidth SlideShape, Width
            Next SlideShape
        End If
    End If
End Sub

Sub ObjectsSwapPosition()
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim Left1, Left2, Top1, Top2 As Single
    
    If ActiveWindow.Selection.ShapeRange.Count = 2 Then
        
        Left1 = ActiveWindow.Selection.ShapeRange(1).left
        Left2 = ActiveWindow.Selection.ShapeRange(2).left
        Top1 = ActiveWindow.Selection.ShapeRange(1).Top
        Top2 = ActiveWindow.Selection.ShapeRange(2).Top
        
        ActiveWindow.Selection.ShapeRange(1).left = Left2
        ActiveWindow.Selection.ShapeRange(2).left = Left1
        ActiveWindow.Selection.ShapeRange(1).Top = Top2
        ActiveWindow.Selection.ShapeRange(2).Top = Top1
        
    ElseIf MyDocument.Selection.HasChildShapeRange Then
        
        If MyDocument.Selection.ChildShapeRange.Count = 2 Then
            
            Left1 = MyDocument.Selection.ChildShapeRange(1).left
            Left2 = MyDocument.Selection.ChildShapeRange(2).left
            Top1 = MyDocument.Selection.ChildShapeRange(1).Top
            Top2 = MyDocument.Selection.ChildShapeRange(2).Top
            
            MyDocument.Selection.ChildShapeRange(1).left = Left2
            MyDocument.Selection.ChildShapeRange(2).left = Left1
            MyDocument.Selection.ChildShapeRange(1).Top = Top2
            MyDocument.Selection.ChildShapeRange(2).Top = Top1
            
        Else
            
            MsgBox "Select two shapes to swap positions."
            
        End If
        
    Else
        
        MsgBox "Select two shapes to swap positions."
        
    End If
    
End Sub

Sub ObjectsSwapPositionCentered()
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim Left1, Left2, Top1, Top2, Width1, Width2, Height1, Height2 As Single
    
    If ActiveWindow.Selection.ShapeRange.Count = 2 Then
        
        Left1 = ActiveWindow.Selection.ShapeRange(1).left
        Left2 = ActiveWindow.Selection.ShapeRange(2).left
        Top1 = ActiveWindow.Selection.ShapeRange(1).Top
        Top2 = ActiveWindow.Selection.ShapeRange(2).Top
        Width1 = ActiveWindow.Selection.ShapeRange(1).Width
        Width2 = ActiveWindow.Selection.ShapeRange(2).Width
        Height1 = ActiveWindow.Selection.ShapeRange(1).Height
        Height2 = ActiveWindow.Selection.ShapeRange(2).Height
        
        ActiveWindow.Selection.ShapeRange(1).left = Left2 + (Width2 - Width1) / 2
        ActiveWindow.Selection.ShapeRange(2).left = Left1 + (Width1 - Width2) / 2
        ActiveWindow.Selection.ShapeRange(1).Top = Top2 + (Height2 - Height1) / 2
        ActiveWindow.Selection.ShapeRange(2).Top = Top1 + (Height1 - Height2) / 2
        
    ElseIf MyDocument.Selection.HasChildShapeRange Then
        
        If MyDocument.Selection.ChildShapeRange.Count = 2 Then
            
            Left1 = MyDocument.Selection.ChildShapeRange(1).left
            Left2 = MyDocument.Selection.ChildShapeRange(2).left
            Top1 = MyDocument.Selection.ChildShapeRange(1).Top
            Top2 = MyDocument.Selection.ChildShapeRange(2).Top
            
            Width1 = ActiveWindow.Selection.ChildShapeRange(1).Width
            Width2 = ActiveWindow.Selection.ChildShapeRange(2).Width
            Height1 = ActiveWindow.Selection.ChildShapeRange(1).Height
            Height2 = ActiveWindow.Selection.ChildShapeRange(2).Height
            
            ActiveWindow.Selection.ChildShapeRange(1).left = Left2 + (Width2 - Width1) / 2
            ActiveWindow.Selection.ChildShapeRange(2).left = Left1 + (Width1 - Width2) / 2
            ActiveWindow.Selection.ChildShapeRange(1).Top = Top2 + (Height2 - Height1) / 2
            ActiveWindow.Selection.ChildShapeRange(2).Top = Top1 + (Height1 - Height2) / 2
            
        Else
            
            MsgBox "Select two shapes to swap positions."
            
        End If
        
    Else
        
        MsgBox "Select two shapes to swap positions."
        
    End If
    
End Sub


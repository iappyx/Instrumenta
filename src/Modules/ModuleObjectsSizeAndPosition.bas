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
    Dim Tallest     As Single
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        Tallest = MyDocument.Selection.ChildShapeRange(1).Height
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            If SlideShape.Height > Tallest Then Tallest = SlideShape.Height
        Next
        
        MyDocument.Selection.ChildShapeRange.Height = Tallest
        
    Else
        Tallest = MyDocument.Selection.ShapeRange(1).Height
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            If SlideShape.Height > Tallest Then Tallest = SlideShape.Height
        Next
        
        MyDocument.Selection.ShapeRange.Height = Tallest
        
    End If
    
End Sub

Sub ObjectsSizeToShortest()
    Set MyDocument = Application.ActiveWindow
    Dim Shortest    As Single
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        Shortest = MyDocument.Selection.ChildShapeRange(1).Height
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            If SlideShape.Height < Shortest Then Shortest = SlideShape.Height
        Next
        
        MyDocument.Selection.ChildShapeRange.Height = Shortest
        
    Else
        
        Shortest = MyDocument.Selection.ShapeRange(1).Height
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            If SlideShape.Height < Shortest Then Shortest = SlideShape.Height
        Next
        
        MyDocument.Selection.ShapeRange.Height = Shortest
        
    End If
    
End Sub

Sub ObjectsSizeToWidest()
    Set MyDocument = Application.ActiveWindow
    Dim Widest      As Single
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        Widest = MyDocument.Selection.ChildShapeRange(1).Width
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            If SlideShape.Width > Widest Then Widest = SlideShape.Width
        Next
        
        MyDocument.Selection.ChildShapeRange.Width = Widest
        
    Else
        Widest = MyDocument.Selection.ShapeRange(1).Width
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            If SlideShape.Width > Widest Then Widest = SlideShape.Width
        Next
        
        MyDocument.Selection.ShapeRange.Width = Widest
        
    End If
    
End Sub

Sub ObjectsSizeToNarrowest()
    Set MyDocument = Application.ActiveWindow
    Dim Narrowest   As Single
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        Narrowest = MyDocument.Selection.ChildShapeRange(1).Width
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            If SlideShape.Width < Narrowest Then Narrowest = SlideShape.Width
        Next
        
        MyDocument.Selection.ChildShapeRange.Width = Narrowest
        
    Else
        
        Narrowest = MyDocument.Selection.ShapeRange(1).Width
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            If SlideShape.Width < Narrowest Then Narrowest = SlideShape.Width
        Next
        
        MyDocument.Selection.ShapeRange.Width = Narrowest
        
    End If
    
End Sub

Sub ObjectsSameHeight()
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        MyDocument.Selection.ChildShapeRange.Height = MyDocument.Selection.ChildShapeRange(1).Height
    Else
        MyDocument.Selection.ShapeRange.Height = MyDocument.Selection.ShapeRange(1).Height
    End If
    
End Sub

Sub ObjectsSameWidth()
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        MyDocument.Selection.ChildShapeRange.Width = MyDocument.Selection.ChildShapeRange(1).Width
    Else
        MyDocument.Selection.ShapeRange.Width = MyDocument.Selection.ShapeRange(1).Width
    End If
    
End Sub

Sub ObjectsSameHeightAndWidth()
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        MyDocument.Selection.ChildShapeRange.Height = MyDocument.Selection.ChildShapeRange(1).Height
        MyDocument.Selection.ChildShapeRange.Width = MyDocument.Selection.ChildShapeRange(1).Width
        
    Else
        MyDocument.Selection.ShapeRange.Height = MyDocument.Selection.ShapeRange(1).Height
        MyDocument.Selection.ShapeRange.Width = MyDocument.Selection.ShapeRange(1).Width
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

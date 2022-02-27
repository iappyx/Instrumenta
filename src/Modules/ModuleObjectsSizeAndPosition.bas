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
    Set myDocument = Application.ActiveWindow
    Dim Tallest     As Single
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.HasChildShapeRange Then
        
        Tallest = myDocument.Selection.ChildShapeRange(1).Height
        
        For Each SlideShape In myDocument.Selection.ChildShapeRange
            If SlideShape.Height > Tallest Then Tallest = SlideShape.Height
        Next
        
        myDocument.Selection.ChildShapeRange.Height = Tallest
        
    Else
        Tallest = myDocument.Selection.ShapeRange(1).Height
        
        For Each SlideShape In myDocument.Selection.ShapeRange
            If SlideShape.Height > Tallest Then Tallest = SlideShape.Height
        Next
        
        myDocument.Selection.ShapeRange.Height = Tallest
        
    End If
    
End Sub

Sub ObjectsSizeToShortest()
    Set myDocument = Application.ActiveWindow
    Dim Shortest    As Single
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.HasChildShapeRange Then
        
        Shortest = myDocument.Selection.ChildShapeRange(1).Height
        
        For Each SlideShape In myDocument.Selection.ChildShapeRange
            If SlideShape.Height < Shortest Then Shortest = SlideShape.Height
        Next
        
        myDocument.Selection.ChildShapeRange.Height = Shortest
        
    Else
        
        Shortest = myDocument.Selection.ShapeRange(1).Height
        
        For Each SlideShape In myDocument.Selection.ShapeRange
            If SlideShape.Height < Shortest Then Shortest = SlideShape.Height
        Next
        
        myDocument.Selection.ShapeRange.Height = Shortest
        
    End If
    
End Sub

Sub ObjectsSizeToWidest()
    Set myDocument = Application.ActiveWindow
    Dim Widest      As Single
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.HasChildShapeRange Then
        
        Widest = myDocument.Selection.ChildShapeRange(1).Width
        
        For Each SlideShape In myDocument.Selection.ChildShapeRange
            If SlideShape.Width > Widest Then Widest = SlideShape.Width
        Next
        
        myDocument.Selection.ChildShapeRange.Width = Widest
        
    Else
        Widest = myDocument.Selection.ShapeRange(1).Width
        
        For Each SlideShape In myDocument.Selection.ShapeRange
            If SlideShape.Width > Widest Then Widest = SlideShape.Width
        Next
        
        myDocument.Selection.ShapeRange.Width = Widest
        
    End If
    
End Sub

Sub ObjectsSizeToNarrowest()
    Set myDocument = Application.ActiveWindow
    Dim Narrowest   As Single
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.HasChildShapeRange Then
        
        Narrowest = myDocument.Selection.ChildShapeRange(1).Width
        
        For Each SlideShape In myDocument.Selection.ChildShapeRange
            If SlideShape.Width < Narrowest Then Narrowest = SlideShape.Width
        Next
        
        myDocument.Selection.ChildShapeRange.Width = Narrowest
        
    Else
        
        Narrowest = myDocument.Selection.ShapeRange(1).Width
        
        For Each SlideShape In myDocument.Selection.ShapeRange
            If SlideShape.Width < Narrowest Then Narrowest = SlideShape.Width
        Next
        
        myDocument.Selection.ShapeRange.Width = Narrowest
        
    End If
    
End Sub

Sub ObjectsSameHeight()
    Set myDocument = Application.ActiveWindow
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.HasChildShapeRange Then
        myDocument.Selection.ChildShapeRange.Height = myDocument.Selection.ChildShapeRange(1).Height
    Else
        myDocument.Selection.ShapeRange.Height = myDocument.Selection.ShapeRange(1).Height
    End If
    
End Sub

Sub ObjectsSameWidth()
    Set myDocument = Application.ActiveWindow
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.HasChildShapeRange Then
        myDocument.Selection.ChildShapeRange.Width = myDocument.Selection.ChildShapeRange(1).Width
    Else
        myDocument.Selection.ShapeRange.Width = myDocument.Selection.ShapeRange(1).Width
    End If
    
End Sub

Sub ObjectsSameHeightAndWidth()
    Set myDocument = Application.ActiveWindow
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If myDocument.Selection.HasChildShapeRange Then
        myDocument.Selection.ChildShapeRange.Height = myDocument.Selection.ChildShapeRange(1).Height
        myDocument.Selection.ChildShapeRange.Width = myDocument.Selection.ChildShapeRange(1).Width
        
    Else
        myDocument.Selection.ShapeRange.Height = myDocument.Selection.ShapeRange(1).Height
        myDocument.Selection.ShapeRange.Width = myDocument.Selection.ShapeRange(1).Width
    End If
    
End Sub

Sub ObjectsSwapPosition()
    Set myDocument = Application.ActiveWindow
    If Not myDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim Left1, Left2, Top1, Top2 As Single
    
    If ActiveWindow.Selection.ShapeRange.Count = 2 Then
        
        Left1 = ActiveWindow.Selection.ShapeRange(1).Left
        Left2 = ActiveWindow.Selection.ShapeRange(2).Left
        Top1 = ActiveWindow.Selection.ShapeRange(1).Top
        Top2 = ActiveWindow.Selection.ShapeRange(2).Top
        
        ActiveWindow.Selection.ShapeRange(1).Left = Left2
        ActiveWindow.Selection.ShapeRange(2).Left = Left1
        ActiveWindow.Selection.ShapeRange(1).Top = Top2
        ActiveWindow.Selection.ShapeRange(2).Top = Top1
        
    ElseIf myDocument.Selection.HasChildShapeRange Then
        
        If myDocument.Selection.ChildShapeRange.Count = 2 Then
            
            Left1 = myDocument.Selection.ChildShapeRange(1).Left
            Left2 = myDocument.Selection.ChildShapeRange(2).Left
            Top1 = myDocument.Selection.ChildShapeRange(1).Top
            Top2 = myDocument.Selection.ChildShapeRange(2).Top
            
            myDocument.Selection.ChildShapeRange(1).Left = Left2
            myDocument.Selection.ChildShapeRange(2).Left = Left1
            myDocument.Selection.ChildShapeRange(1).Top = Top2
            myDocument.Selection.ChildShapeRange(2).Top = Top1
            
        Else
            
            MsgBox "Select two shapes to swap positions."
            
        End If
        
    Else
        
        MsgBox "Select two shapes to swap positions."
        
    End If
    
End Sub

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
    Tallest = myDocument.Selection.ShapeRange(1).Height
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        If SlideShape.Height > Tallest Then Tallest = SlideShape.Height
    Next
    
    myDocument.Selection.ShapeRange.Height = Tallest
    
End Sub

Sub ObjectsSizeToShortest()
    Set myDocument = Application.ActiveWindow
    Dim Shortest    As Single
    Shortest = myDocument.Selection.ShapeRange(1).Height
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        If SlideShape.Height < Shortest Then Shortest = SlideShape.Height
    Next
    
    myDocument.Selection.ShapeRange.Height = Shortest
    
End Sub

Sub ObjectsSizeToWidest()
    Set myDocument = Application.ActiveWindow
    Dim Widest      As Single
    Widest = myDocument.Selection.ShapeRange(1).Width
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        If SlideShape.Width > Widest Then Widest = SlideShape.Width
    Next
    
    myDocument.Selection.ShapeRange.Width = Widest
    
End Sub

Sub ObjectsSizeToNarrowest()
    Set myDocument = Application.ActiveWindow
    Dim Narrowest   As Single
    Narrowest = myDocument.Selection.ShapeRange(1).Width
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        If SlideShape.Width < Narrowest Then Narrowest = SlideShape.Width
    Next
    
    myDocument.Selection.ShapeRange.Width = Narrowest
    
End Sub

Sub ObjectsSameHeight()
    Set myDocument = Application.ActiveWindow
    
    myDocument.Selection.ShapeRange.Height = myDocument.Selection.ShapeRange(1).Height
    
End Sub

Sub ObjectsSameWidth()
    Set myDocument = Application.ActiveWindow
    
    myDocument.Selection.ShapeRange.Width = myDocument.Selection.ShapeRange(1).Width
    
End Sub

Sub ObjectsSameHeightAndWidth()
    Set myDocument = Application.ActiveWindow
    
    myDocument.Selection.ShapeRange.Height = myDocument.Selection.ShapeRange(1).Height
    myDocument.Selection.ShapeRange.Width = myDocument.Selection.ShapeRange(1).Width
    
End Sub

Sub ObjectsSwapPosition()
    
    If ActiveWindow.Selection.ShapeRange.Count = 2 Then
    
    Dim Left1, Left2, Top1, Top2 As Single
    
    Left1 = ActiveWindow.Selection.ShapeRange(1).Left
    Left2 = ActiveWindow.Selection.ShapeRange(2).Left
    Top1 = ActiveWindow.Selection.ShapeRange(1).Top
    Top2 = ActiveWindow.Selection.ShapeRange(2).Top
    
    ActiveWindow.Selection.ShapeRange(1).Left = Left2
    ActiveWindow.Selection.ShapeRange(2).Left = Left1
    ActiveWindow.Selection.ShapeRange(1).Top = Top2
    ActiveWindow.Selection.ShapeRange(2).Top = Top1
    
    Else
    
    MsgBox "Select two shapes to swap positions."
    
    End If
    
End Sub

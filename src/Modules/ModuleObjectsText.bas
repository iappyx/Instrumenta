Attribute VB_Name = "ModuleObjectsText"
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

Sub ObjectsRemoveText()
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
    
    myDocument.Selection.ShapeRange.TextFrame.TextRange.Text = ""
    
    Else
    
    MsgBox "No text capable shape selected."
    
    End If
    
End Sub

Sub ObjectsSwapTextNoFormatting()

    Dim text1, text2 As String
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange.Count = 2 Then
    
    If myDocument.Selection.ShapeRange(1).HasTextFrame And myDocument.Selection.ShapeRange(2).HasTextFrame Then
    
    text1 = myDocument.Selection.ShapeRange(1).TextFrame.TextRange.Text
    text2 = myDocument.Selection.ShapeRange(2).TextFrame.TextRange.Text
    myDocument.Selection.ShapeRange(1).TextFrame.TextRange.Text = text2
    myDocument.Selection.ShapeRange(2).TextFrame.TextRange.Text = text1
    
    Else
    
    MsgBox "Select two shapes that (can) have text."
    
    End If
    
    
    Else
    
    MsgBox "Select two shapes to swap their text."
    
    End If
    

End Sub

Sub ObjectsSwapText()

    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange.Count = 2 Then
    
    If myDocument.Selection.ShapeRange(1).HasTextFrame And myDocument.Selection.ShapeRange(2).HasTextFrame Then
    
    Dim SlidePlaceHolder As PowerPoint.Shape
    Set SlidePlaceHolder = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=100, Height:=100)
    
    myDocument.Selection.ShapeRange(1).TextFrame.TextRange.Cut
    SlidePlaceHolder.TextFrame.TextRange.Paste
    
    myDocument.Selection.ShapeRange(2).TextFrame.TextRange.Cut
    myDocument.Selection.ShapeRange(1).TextFrame.TextRange.Paste
    
    SlidePlaceHolder.TextFrame.TextRange.Cut
    myDocument.Selection.ShapeRange(2).TextFrame.TextRange.Paste
       
    SlidePlaceHolder.Delete
    
    Else
    
    MsgBox "Select two shapes that (can) have text."
    
    End If
    
    
    Else
    
    MsgBox "Select two shapes to swap their text."
    
    End If

End Sub

Sub ObjectsMarginsToZero()
    
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
    
        With myDocument.Selection.ShapeRange.TextFrame
        .MarginBottom = 0
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        
    End With
    
    Else
    
    MsgBox "No text capable shape selected."
    
    End If
    
End Sub

Sub ObjectsMarginsIncrease()
    
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
    
    With myDocument.Selection.ShapeRange.TextFrame
        .MarginBottom = .MarginBottom + 0.2
        .MarginLeft = .MarginLeft + 0.2
        .MarginRight = .MarginRight + 0.2
        .MarginTop = .MarginTop + 0.2
        
    End With
    
    Else
    
    MsgBox "No text capable shape selected."
    
    End If
    
End Sub

Sub ObjectsMarginsDecrease()
    
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
    
    With myDocument.Selection.ShapeRange.TextFrame
        If .MarginBottom >= 0.2 Then
            .MarginBottom = .MarginBottom - 0.2
        End If
        If .MarginLeft >= 0.2 Then
            .MarginLeft = .MarginLeft - 0.2
        End If
        If .MarginRight >= 0.2 Then
            .MarginRight = .MarginRight - 0.2
        End If
        If .MarginTop >= 0.2 Then
            .MarginTop = .MarginTop - 0.2
        End If
        
    End With
    
    Else
    
    MsgBox "No text capable shape selected."
    
    End If
    
End Sub

Sub ObjectsTextWordwrapToggle()
    
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
     
    If myDocument.Selection.ShapeRange.TextFrame.WordWrap = True Then
        myDocument.Selection.ShapeRange.TextFrame.WordWrap = False
    Else
        myDocument.Selection.ShapeRange.TextFrame.WordWrap = True
    End If
    
    Else
    
    MsgBox "No text capable shape selected."
    
    End If
    
End Sub

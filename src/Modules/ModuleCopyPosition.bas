Attribute VB_Name = "ModuleCopyPosition"
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

Public TopToCopy, LeftToCopy, WidthToCopy, HeightToCopy As Long
Public PositionCopied As Boolean

Sub CopyPosition()
    
    Set myDocument = Application.ActiveWindow
    
    TopToCopy = myDocument.Selection.ShapeRange(1).Top
    LeftToCopy = myDocument.Selection.ShapeRange(1).Left
    WidthToCopy = myDocument.Selection.ShapeRange(1).Width
    HeightToCopy = myDocument.Selection.ShapeRange(1).Height
    PositionCopied = True
    
End Sub

Sub PastePosition()
    
    If PositionCopied = True Then
        Set myDocument = Application.ActiveWindow
        myDocument.Selection.ShapeRange(1).Top = TopToCopy
        myDocument.Selection.ShapeRange(1).Left = LeftToCopy
    Else
        MsgBox "No dimensions available. First copy position / dimension of a shape."
    End If
    
End Sub

Sub PastePositionAndDimensions()
    
    If PositionCopied = True Then
        Set myDocument = Application.ActiveWindow
        myDocument.Selection.ShapeRange(1).Top = TopToCopy
        myDocument.Selection.ShapeRange(1).Left = LeftToCopy
        myDocument.Selection.ShapeRange(1).Width = WidthToCopy
        myDocument.Selection.ShapeRange(1).Height = HeightToCopy
    Else
        MsgBox "No dimensions available. First copy position / dimension of a shape."
    End If
    
End Sub

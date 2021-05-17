Attribute VB_Name = "ModuleObjectsRoundedCorners"
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

Sub ObjectsCopyRoundedCorner()
    Dim SlideShape  As PowerPoint.Shape
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim ShapeRadius As Single
    If Application.ActiveWindow.Selection.ShapeRange(1).Adjustments.Count > 0 Then
    
    ShapeRadius = myDocument.Selection.ShapeRange(1).Adjustments(1) / (1 / (myDocument.Selection.ShapeRange(1).Height + myDocument.Selection.ShapeRange(1).Width))
    
    If myDocument.Selection.ShapeRange(1).Adjustments.Count > 1 Then
        ShapeRadius2 = myDocument.Selection.ShapeRange(1).Adjustments(2) / (1 / (myDocument.Selection.ShapeRange(1).Height + myDocument.Selection.ShapeRange(1).Width))
    End If
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        With SlideShape
            .AutoShapeType = myDocument.Selection.ShapeRange(1).AutoShapeType
            .Adjustments(1) = (1 / (SlideShape.Height + SlideShape.Width)) * ShapeRadius
            If myDocument.Selection.ShapeRange(1).Adjustments.Count > 1 Then
                .Adjustments(2) = (1 / (SlideShape.Height + SlideShape.Width)) * ShapeRadius2
            End If
        End With
    Next
    
    End If
    
    End If
    
End Sub

Sub ObjectsCopyShapeTypeAndAdjustments()
    Dim SlideShape  As PowerPoint.Shape
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim AdjustmentsCount As Long
    Dim ShapeCount  As Long
    
    For ShapeCount = 2 To ActiveWindow.Selection.ShapeRange.Count
        
        myDocument.Selection.ShapeRange(ShapeCount).AutoShapeType = myDocument.Selection.ShapeRange(1).AutoShapeType
        
        For AdjustmentsCount = 1 To myDocument.Selection.ShapeRange(1).Adjustments.Count
            
            myDocument.Selection.ShapeRange(ShapeCount).Adjustments(AdjustmentsCount) = myDocument.Selection.ShapeRange(1).Adjustments(AdjustmentsCount)
            
        Next AdjustmentsCount
        
    Next ShapeCount
    
    End If
    
End Sub

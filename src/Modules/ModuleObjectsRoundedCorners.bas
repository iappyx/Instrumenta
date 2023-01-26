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
    Dim SlideShape  As PowerPoint.shape
    Set MyDocument = Application.ActiveWindow
    Dim ShapeRadius As Single
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        
    ElseIf MyDocument.Selection.HasChildShapeRange Then
        
        If Application.ActiveWindow.Selection.ChildShapeRange(1).Adjustments.Count > 0 Then
            
            ShapeRadius = MyDocument.Selection.ChildShapeRange(1).Adjustments(1) / (1 / (MyDocument.Selection.ChildShapeRange(1).Height + MyDocument.Selection.ChildShapeRange(1).Width))
            
            If MyDocument.Selection.ChildShapeRange(1).Adjustments.Count > 1 Then
                ShapeRadius2 = MyDocument.Selection.ChildShapeRange(1).Adjustments(2) / (1 / (MyDocument.Selection.ChildShapeRange(1).Height + MyDocument.Selection.ChildShapeRange(1).Width))
            End If
            
            For Each SlideShape In ActiveWindow.Selection.ChildShapeRange
                With SlideShape
                    .AutoShapeType = MyDocument.Selection.ChildShapeRange(1).AutoShapeType
                    .Adjustments(1) = (1 / (SlideShape.Height + SlideShape.Width)) * ShapeRadius
                    If MyDocument.Selection.ChildShapeRange(1).Adjustments.Count > 1 Then
                        .Adjustments(2) = (1 / (SlideShape.Height + SlideShape.Width)) * ShapeRadius2
                    End If
                End With
            Next
            
        End If
        
    Else
        
        For i = 1 To Application.ActiveWindow.Selection.ShapeRange.Count
        
            If Application.ActiveWindow.Selection.ShapeRange(i).Type = msoGroup Then
                MsgBox "One of the selected shapes is a group."
                Exit Sub
            End If
                
        Next i
        
        
        If Application.ActiveWindow.Selection.ShapeRange(1).Adjustments.Count > 0 Then
            
            ShapeRadius = MyDocument.Selection.ShapeRange(1).Adjustments(1) / (1 / (MyDocument.Selection.ShapeRange(1).Height + MyDocument.Selection.ShapeRange(1).Width))
            
            If MyDocument.Selection.ShapeRange(1).Adjustments.Count > 1 Then
                ShapeRadius2 = MyDocument.Selection.ShapeRange(1).Adjustments(2) / (1 / (MyDocument.Selection.ShapeRange(1).Height + MyDocument.Selection.ShapeRange(1).Width))
            End If
            
            For Each SlideShape In ActiveWindow.Selection.ShapeRange
                With SlideShape
                    .AutoShapeType = MyDocument.Selection.ShapeRange(1).AutoShapeType
                    .Adjustments(1) = (1 / (SlideShape.Height + SlideShape.Width)) * ShapeRadius
                    If MyDocument.Selection.ShapeRange(1).Adjustments.Count > 1 Then
                        .Adjustments(2) = (1 / (SlideShape.Height + SlideShape.Width)) * ShapeRadius2
                    End If
                End With
            Next
            
        End If
        
    End If
    
End Sub

Sub ObjectsCopyShapeTypeAndAdjustments()
    Dim SlideShape  As PowerPoint.shape
    Set MyDocument = Application.ActiveWindow
    Dim AdjustmentsCount As Long
    Dim ShapeCount  As Long
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        
    ElseIf MyDocument.Selection.HasChildShapeRange Then
        
        For ShapeCount = 2 To ActiveWindow.Selection.ChildShapeRange.Count
            
            MyDocument.Selection.ChildShapeRange(ShapeCount).AutoShapeType = MyDocument.Selection.ChildShapeRange(1).AutoShapeType
            
            For AdjustmentsCount = 1 To MyDocument.Selection.ChildShapeRange(1).Adjustments.Count
                
                MyDocument.Selection.ChildShapeRange(ShapeCount).Adjustments(AdjustmentsCount) = MyDocument.Selection.ChildShapeRange(1).Adjustments(AdjustmentsCount)
                
            Next AdjustmentsCount
            
        Next ShapeCount
        
    Else
        
        For i = 1 To Application.ActiveWindow.Selection.ShapeRange.Count
        
            If Application.ActiveWindow.Selection.ShapeRange(i).Type = msoGroup Then
                MsgBox "One of the selected shapes is a group."
                Exit Sub
            End If
                
        Next i
        
        For ShapeCount = 2 To ActiveWindow.Selection.ShapeRange.Count
            
            MyDocument.Selection.ShapeRange(ShapeCount).AutoShapeType = MyDocument.Selection.ShapeRange(1).AutoShapeType
            
            For AdjustmentsCount = 1 To MyDocument.Selection.ShapeRange(1).Adjustments.Count
                
                MyDocument.Selection.ShapeRange(ShapeCount).Adjustments(AdjustmentsCount) = MyDocument.Selection.ShapeRange(1).Adjustments(AdjustmentsCount)
                
            Next AdjustmentsCount
            
        Next ShapeCount
        
    End If
    
End Sub

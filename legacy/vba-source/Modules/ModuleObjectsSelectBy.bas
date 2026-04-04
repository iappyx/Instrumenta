Attribute VB_Name = "ModuleObjectsSelectBy"
'MIT License

'Copyright (c) 2021 - 2026 iappyx

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

Sub ObjectsSelectBySameType()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.shape
    Dim selectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve selectedShapes(0)
    selectedShapes(0) = SlideShape.name
    
    For Each SlideShapeToCheck In MyDocument.View.Slide.shapes
        
        If (SlideShapeToCheck.AutoShapeType = SlideShape.AutoShapeType) Then
            
            If (SlideShapeToCheck.name <> SlideShape.name) Then
                ReDim Preserve selectedShapes(ShapeCount + 1)
                selectedShapes(ShapeCount) = SlideShapeToCheck.name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    MyDocument.View.Slide.shapes.Range(selectedShapes).Select
    
    End If
    
End Sub

Sub ObjectsSelectBySameFillAndLineColor()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.shape
    Dim selectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve selectedShapes(0)
    selectedShapes(0) = SlideShape.name
    
    For Each SlideShapeToCheck In MyDocument.View.Slide.shapes
        
        If (SlideShapeToCheck.Fill.ForeColor.RGB = SlideShape.Fill.ForeColor.RGB) And (SlideShapeToCheck.line.ForeColor.RGB = SlideShape.line.ForeColor.RGB) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.visible = True) And (SlideShapeToCheck.line.visible = True) Then
            
            If (SlideShapeToCheck.name <> SlideShape.name) Then
                ReDim Preserve selectedShapes(ShapeCount + 1)
                selectedShapes(ShapeCount) = SlideShapeToCheck.name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    MyDocument.View.Slide.shapes.Range(selectedShapes).Select
    
    End If
End Sub

Sub ObjectsSelectBySameFillColor()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.shape
    Dim selectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve selectedShapes(0)
    selectedShapes(0) = SlideShape.name
    
    For Each SlideShapeToCheck In MyDocument.View.Slide.shapes
        
        If (SlideShapeToCheck.Fill.ForeColor.RGB = SlideShape.Fill.ForeColor.RGB) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.visible = True) Then
            
            If (SlideShapeToCheck.name <> SlideShape.name) Then
                ReDim Preserve selectedShapes(ShapeCount + 1)
                selectedShapes(ShapeCount) = SlideShapeToCheck.name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    MyDocument.View.Slide.shapes.Range(selectedShapes).Select
    
    End If
    
End Sub

Sub ObjectsSelectBySameLineColor()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.shape
    Dim selectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve selectedShapes(0)
    selectedShapes(0) = SlideShape.name
    
    For Each SlideShapeToCheck In MyDocument.View.Slide.shapes
        
        If (SlideShapeToCheck.line.ForeColor.RGB = SlideShape.line.ForeColor.RGB) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.line.visible = True) Then
            
            If (SlideShapeToCheck.name <> SlideShape.name) Then
                ReDim Preserve selectedShapes(ShapeCount + 1)
                selectedShapes(ShapeCount) = SlideShapeToCheck.name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    MyDocument.View.Slide.shapes.Range(selectedShapes).Select
    
    End If
End Sub

Sub ObjectsSelectBySameWidthAndHeight()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.shape
    Dim selectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve selectedShapes(0)
    selectedShapes(0) = SlideShape.name
    
    For Each SlideShapeToCheck In MyDocument.View.Slide.shapes
        
        If (SlideShapeToCheck.width = SlideShape.width) And (SlideShapeToCheck.height = SlideShape.height) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.visible = True) Then
            
            If (SlideShapeToCheck.name <> SlideShape.name) Then
                ReDim Preserve selectedShapes(ShapeCount + 1)
                selectedShapes(ShapeCount) = SlideShapeToCheck.name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    MyDocument.View.Slide.shapes.Range(selectedShapes).Select
    
    End If
End Sub

Sub ObjectsSelectBySameWidth()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.shape
    Dim selectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve selectedShapes(0)
    selectedShapes(0) = SlideShape.name
    
    For Each SlideShapeToCheck In MyDocument.View.Slide.shapes
        
        If (SlideShapeToCheck.width = SlideShape.width) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.visible = True) Then
            
            If (SlideShapeToCheck.name <> SlideShape.name) Then
                ReDim Preserve selectedShapes(ShapeCount + 1)
                selectedShapes(ShapeCount) = SlideShapeToCheck.name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    MyDocument.View.Slide.shapes.Range(selectedShapes).Select
    
    End If
End Sub

Sub ObjectsSelectBySameHeight()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.shape
    Dim selectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve selectedShapes(0)
    selectedShapes(0) = SlideShape.name
    
    For Each SlideShapeToCheck In MyDocument.View.Slide.shapes
        
        If (SlideShapeToCheck.height = SlideShape.height) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.visible = True) Then
            
            If (SlideShapeToCheck.name <> SlideShape.name) Then
                ReDim Preserve selectedShapes(ShapeCount + 1)
                selectedShapes(ShapeCount) = SlideShapeToCheck.name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    MyDocument.View.Slide.shapes.Range(selectedShapes).Select
    
    End If
End Sub

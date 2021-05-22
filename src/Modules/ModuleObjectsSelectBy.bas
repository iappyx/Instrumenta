Attribute VB_Name = "ModuleObjectsSelectBy"
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

Sub ObjectsSelectBySameType()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.AutoShapeType = SlideShape.AutoShapeType) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
    End If
    
End Sub

Sub ObjectsSelectBySameFillAndLineColor()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Fill.ForeColor.RGB = SlideShape.Fill.ForeColor.RGB) And (SlideShapeToCheck.Line.ForeColor.RGB = SlideShape.Line.ForeColor.RGB) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.Visible = True) And (SlideShapeToCheck.Line.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
    End If
End Sub

Sub ObjectsSelectBySameFillColor()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Fill.ForeColor.RGB = SlideShape.Fill.ForeColor.RGB) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
    End If
    
End Sub

Sub ObjectsSelectBySameLineColor()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Line.ForeColor.RGB = SlideShape.Line.ForeColor.RGB) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Line.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
    End If
End Sub

Sub ObjectsSelectBySameWidthAndHeight()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Width = SlideShape.Width) And (SlideShapeToCheck.Height = SlideShape.Height) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
    End If
End Sub

Sub ObjectsSelectBySameWidth()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Width = SlideShape.Width) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
    End If
End Sub

Sub ObjectsSelectBySameHeight()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Height = SlideShape.Height) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
    End If
End Sub

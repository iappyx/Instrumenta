Attribute VB_Name = "ModuleObjectsRotateAndFlip"
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


Sub ObjectsFlipHorizontal()
    Set MyDocument = Application.ActiveWindow
    
    Dim shp As shape
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        For Each shp In MyDocument.Selection.ChildShapeRange
            If shp.Type <> msoTable Then shp.Flip msoFlipHorizontal
        Next shp
        
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        If MyDocument.Selection.ShapeRange(1).Type <> msoTable Then
            MyDocument.Selection.ShapeRange(1).Flip msoFlipHorizontal
        End If
        
    Else
        For Each shp In MyDocument.Selection.ShapeRange
            If shp.Type <> msoTable Then shp.Flip msoFlipHorizontal
        Next shp
    End If
End Sub

Sub ObjectsFlipVertical()
    Set MyDocument = Application.ActiveWindow
    
    Dim shp As shape
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        For Each shp In MyDocument.Selection.ChildShapeRange
            If shp.Type <> msoTable Then shp.Flip msoFlipVertical
        Next shp
        
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        If MyDocument.Selection.ShapeRange(1).Type <> msoTable Then
            MyDocument.Selection.ShapeRange(1).Flip msoFlipVertical
        End If
        
    Else
        For Each shp In MyDocument.Selection.ShapeRange
            If shp.Type <> msoTable Then shp.Flip msoFlipVertical
        Next shp
    End If
End Sub

Sub ObjectsRotateClockwise()
    Set MyDocument = Application.ActiveWindow
    
    Dim shp As shape
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        For Each shp In MyDocument.Selection.ChildShapeRange
            If shp.Type <> msoTable Then shp.rotation = shp.rotation + 90
        Next shp
        
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        If MyDocument.Selection.ShapeRange(1).Type <> msoTable Then
            MyDocument.Selection.ShapeRange(1).rotation = MyDocument.Selection.ShapeRange(1).rotation + 90
        End If
        
    Else
        For Each shp In MyDocument.Selection.ShapeRange
            If shp.Type <> msoTable Then shp.rotation = shp.rotation + 90
        Next shp
    End If
End Sub

Sub ObjectsRotateCounterclockwise()

    Set MyDocument = Application.ActiveWindow
    
    Dim shp As shape
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        For Each shp In MyDocument.Selection.ChildShapeRange
            If shp.Type <> msoTable Then shp.rotation = shp.rotation - 90
        Next shp
        
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        If MyDocument.Selection.ShapeRange(1).Type <> msoTable Then
            MyDocument.Selection.ShapeRange(1).rotation = MyDocument.Selection.ShapeRange(1).rotation - 90
        End If
        
    Else
        For Each shp In MyDocument.Selection.ShapeRange
            If shp.Type <> msoTable Then shp.rotation = shp.rotation - 90
        Next shp
    End If
End Sub


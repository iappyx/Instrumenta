Attribute VB_Name = "ModuleObjectsRowsAndColumns"
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

Sub GroupShapesByColumns()
    Dim SlideShape         As shape
    Dim ShapeGroups      As Collection
    Set ShapeGroups = New Collection
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
    SlideShape.name = "Shape " & SlideShape.id
    Next SlideShape
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        
        Dim ShapeShapeGroupExists As Boolean
        ShapeShapeGroupExists = False
        
        For Each ShapeGroup In ShapeGroups
            
            If (SlideShape.left + SlideShape.width) >= ShapeGroup(1).left And SlideShape.left <= (ShapeGroup(1).left + ShapeGroup(1).width) Then
                ShapeGroup.Add SlideShape
                ShapeShapeGroupExists = True
                Exit For
            End If
        Next ShapeGroup
        
        If Not ShapeShapeGroupExists Then
            Set NewShapeGroup = New Collection
            NewShapeGroup.Add SlideShape
            ShapeGroups.Add NewShapeGroup
        End If
    Next SlideShape
    
    For Each ShapeGroup In ShapeGroups
        Dim ShapeNames() As String
        ReDim ShapeNames(ShapeGroup.count - 1)
        For i = 1 To ShapeGroup.count
            ShapeNames(i - 1) = ShapeGroup(i).name
        Next i
        
        If ShapeGroup.count > 1 Then
            ActiveWindow.Selection.SlideRange.shapes.Range(ShapeNames).Group
        End If
    Next ShapeGroup
    
End Sub

Sub GroupShapesByRows()
    Dim SlideShape         As shape
    Dim ShapeGroups      As Collection
    Set ShapeGroups = New Collection
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
    SlideShape.name = "Shape " & SlideShape.id
    Next SlideShape
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        
        Dim ShapeShapeGroupExists As Boolean
        ShapeShapeGroupExists = False
        
        For Each ShapeGroup In ShapeGroups
            
            If (SlideShape.Top + SlideShape.height) >= ShapeGroup(1).Top And SlideShape.Top <= (ShapeGroup(1).Top + ShapeGroup(1).height) Then
                ShapeGroup.Add SlideShape
                ShapeShapeGroupExists = True
                Exit For
            End If
        Next ShapeGroup
        
        If Not ShapeShapeGroupExists Then
            Set NewShapeGroup = New Collection
            NewShapeGroup.Add SlideShape
            ShapeGroups.Add NewShapeGroup
        End If
    Next SlideShape
    
    For Each ShapeGroup In ShapeGroups
        Dim ShapeNames() As String
        ReDim ShapeNames(ShapeGroup.count - 1)
        For i = 1 To ShapeGroup.count
            ShapeNames(i - 1) = ShapeGroup(i).name
        Next i
        
        If ShapeGroup.count > 1 Then
            ActiveWindow.Selection.SlideRange.shapes.Range(ShapeNames).Group
        End If
    Next ShapeGroup
    
End Sub


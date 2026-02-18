Attribute VB_Name = "ModuleMaster"
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

Sub MoveSelectedShapeToMaster()
    ShapeToMasterSlide True, False
End Sub

Sub CopySelectedShapeToMaster()
    ShapeToMasterSlide False, False
End Sub

Sub MoveSelectedShapeToAllMasters()
    ShapeToMasterSlide True, True
End Sub

Sub CopySelectedShapeToAllMasters()
    ShapeToMasterSlide False, True
End Sub

Sub MoveSelectedShapeToUsedMasters()
    ShapeToMasterSlide True, True, True
End Sub

Sub CopySelectedShapeToUsedMasters()
    ShapeToMasterSlide False, True, True
End Sub


Sub ShapeToMasterSlide(deleteOriginal As Boolean, toAllLayouts As Boolean, Optional onlyUsedLayouts As Boolean = False)

    Dim shp As shape
    Dim dup As shape
    Dim sld As Slide
    Dim layout As CustomLayout

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Select a single shape first."
        Exit Sub
    End If

    If ActiveWindow.Selection.ShapeRange.count <> 1 Then
        MsgBox "Select exactly one shape."
        Exit Sub
    End If

    Set shp = ActiveWindow.Selection.ShapeRange(1)
    Set sld = ActiveWindow.View.Slide

    shp.Copy

    If toAllLayouts Then

        For Each layout In sld.master.CustomLayouts

            If (Not onlyUsedLayouts) Or LayoutIsUsed(layout) Then

                layout.Shapes.Paste
                Set dup = layout.Shapes(layout.Shapes.count)

                dup.left = shp.left
                dup.Top = shp.Top
                dup.width = shp.width
                dup.height = shp.height

            End If

        Next layout

    Else

        Set layout = sld.CustomLayout

        layout.Shapes.Paste
        Set dup = layout.Shapes(layout.Shapes.count)

        dup.left = shp.left
        dup.Top = shp.Top
        dup.width = shp.width
        dup.height = shp.height

    End If

    If deleteOriginal Then
        shp.Delete
        If toAllLayouts Then
            If onlyUsedLayouts Then
                MsgBox "Shape moved to all used master layouts."
            Else
                MsgBox "Shape moved to all master layouts."
            End If
        Else
            MsgBox "Shape moved to the current master layout."
        End If
    Else
        If toAllLayouts Then
            If onlyUsedLayouts Then
                MsgBox "Shape copied to all used master layouts."
            Else
                MsgBox "Shape copied to all master layouts."
            End If
        Else
            MsgBox "Shape copied to the current master layout."
        End If
    End If

End Sub


Function LayoutIsUsed(layout As CustomLayout) As Boolean
    Dim s As Slide
    For Each s In ActivePresentation.Slides
        If s.CustomLayout.index = layout.index Then
            LayoutIsUsed = True
            Exit Function
        End If
    Next s
End Function

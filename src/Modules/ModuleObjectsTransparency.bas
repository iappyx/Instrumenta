Attribute VB_Name = "ModuleObjectsTransparency"
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

Sub IncreaseShapeTransparency()
    Dim shp, grpShp As shape
    Set MyDocument = Application.ActiveWindow
    
    If MyDocument.Selection.Type = ppSelectionShapes Then
        If MyDocument.Selection.HasChildShapeRange Then
            For Each shp In MyDocument.Selection.ChildShapeRange
                If shp.Type = msoGroup Then
                    For Each grpShp In shp.GroupItems
                        SetTransparency grpShp
                    Next grpShp
                Else
                    SetTransparency shp
                End If
            Next shp
        Else
            For Each shp In MyDocument.Selection.ShapeRange

                If shp.Type = msoGroup Then
                    For Each grpShp In shp.GroupItems
                        SetTransparency grpShp
                    Next grpShp
                Else
                    SetTransparency shp
                End If
            Next shp
        End If
    Else
        MsgBox "Please select one or more shapes."
    End If
End Sub


Sub SetTransparency(ByVal shp As shape)
    Dim currentFillTransparency As Single
    Dim adjustedTransparency As Single
    
    If shp.Fill.Type = msoFillSolid Then
        currentFillTransparency = shp.Fill.Transparency
        adjustedTransparency = currentFillTransparency + 0.1

        If adjustedTransparency > 1 Then
            adjustedTransparency = 0
        End If

        shp.Fill.Transparency = adjustedTransparency
        shp.Line.Transparency = adjustedTransparency
    End If
    
End Sub


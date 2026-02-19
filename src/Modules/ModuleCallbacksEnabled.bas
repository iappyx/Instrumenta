Attribute VB_Name = "ModuleCallbacksEnabled"
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

Sub EnableWhenShapesSelected(control As IRibbonControl, ByRef enabled)
    
If Not SettingContextualButtons Then enabled = True: Exit Sub

    On Error Resume Next
    enabled = (ActiveWindow.Selection.Type = ppSelectionShapes)
End Sub


Sub EnableWhenSlidesSelected(control As IRibbonControl, ByRef enabled)

If Not SettingContextualButtons Then enabled = True: Exit Sub

    On Error Resume Next
    enabled = (ActiveWindow.Selection.Type = ppSelectionSlides)
End Sub

Sub EnableWhenTextSelected(control As IRibbonControl, ByRef enabled)
If Not SettingContextualButtons Then enabled = True: Exit Sub

    On Error Resume Next
    enabled = (ActiveWindow.Selection.Type = ppSelectionText)
End Sub

Sub EnableWhenMultipleShapesSelected(control As IRibbonControl, ByRef enabled)
If Not SettingContextualButtons Then enabled = True: Exit Sub

    On Error Resume Next

    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        enabled = (ActiveWindow.Selection.ShapeRange.count >= 2)
    Else
        enabled = False
    End If
End Sub

Sub EnableWhenShapeOrText(control As IRibbonControl, ByRef enabled)
If Not SettingContextualButtons Then enabled = True: Exit Sub

    On Error Resume Next

    Select Case ActiveWindow.Selection.Type
        Case ppSelectionShapes, ppSelectionText
            enabled = True
        Case Else
            enabled = False
    End Select
End Sub

Sub EnableWhenExactlyOneShape(control As IRibbonControl, ByRef enabled)

If Not SettingContextualButtons Then enabled = True: Exit Sub

    On Error Resume Next

    With ActiveWindow.Selection
        enabled = (.Type = ppSelectionShapes And .ShapeRange.count = 1)
    End With
End Sub

Public Sub EnableWhenInTable(control As IRibbonControl, ByRef enabled)

If Not SettingContextualButtons Then enabled = True: Exit Sub

    enabled = False
    
    Dim sel As Selection
    On Error GoTo Done
    Set sel = ActiveWindow.Selection
    
    If sel.Type = ppSelectionShapes Then
        If sel.ShapeRange.count = 1 Then
            If sel.ShapeRange(1).HasTable Then
                enabled = True
            End If
        End If
        GoTo Done
    End If
    
    If sel.Type = ppSelectionText Then
        Dim oShp As shape
        Set oShp = sel.ShapeRange(1)
        If oShp.HasTable Then
            enabled = True
        End If
    End If

Done:
    On Error GoTo 0
End Sub


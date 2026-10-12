Attribute VB_Name = "ModuleCallbacksEnabled"
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

Private Const SEL_SHAPES As Long = 1
Private Const SEL_ONE_SHAPE As Long = 2
Private Const SEL_MULTIPLE_SHAPES As Long = 4
Private Const SEL_TEXT As Long = 8
Private Const SEL_SLIDES As Long = 16
Private Const SEL_IN_TABLE As Long = 32

Private CachedSelectionState As Long
Private SelectionStateKnown As Boolean

'Reads the selection once and reduces it to the few facts the getEnabled callbacks need.
'The ribbon only has to be invalidated when this state changes, not on every click.
Private Function ComputeSelectionState(sel As Selection) As Long
    Dim state As Long
    On Error GoTo Done

    Select Case sel.Type
        Case ppSelectionShapes
            state = SEL_SHAPES
            Select Case sel.ShapeRange.count
                Case 1
                    state = state Or SEL_ONE_SHAPE
                    If sel.ShapeRange(1).HasTable Then state = state Or SEL_IN_TABLE
                Case Is >= 2
                    state = state Or SEL_MULTIPLE_SHAPES
            End Select
        Case ppSelectionText
            state = SEL_TEXT
            If sel.ShapeRange(1).HasTable Then state = state Or SEL_IN_TABLE
        Case ppSelectionSlides
            state = SEL_SLIDES
    End Select

Done:
    ComputeSelectionState = state
End Function

'Returns True when the selection state differs from the cached one
Public Function RefreshSelectionState(Optional sel As Selection) As Boolean
    Dim state As Long

    If sel Is Nothing Then
        On Error Resume Next
        Set sel = ActiveWindow.Selection
        On Error GoTo 0
    End If

    If Not sel Is Nothing Then state = ComputeSelectionState(sel)

    RefreshSelectionState = (Not SelectionStateKnown) Or (state <> CachedSelectionState)
    CachedSelectionState = state
    SelectionStateKnown = True
End Function

Private Function SelectionIs(flags As Long) As Boolean
    If Not SelectionStateKnown Then RefreshSelectionState
    SelectionIs = ((CachedSelectionState And flags) <> 0)
End Function

Sub EnableWhenShapesSelected(control As IRibbonControl, ByRef enabled)
If Not SettingContextualButtons Then enabled = True: Exit Sub
    enabled = SelectionIs(SEL_SHAPES)
End Sub

Sub EnableWhenSlidesSelected(control As IRibbonControl, ByRef enabled)
If Not SettingContextualButtons Then enabled = True: Exit Sub
    enabled = SelectionIs(SEL_SLIDES)
End Sub

Sub EnableWhenTextSelected(control As IRibbonControl, ByRef enabled)
If Not SettingContextualButtons Then enabled = True: Exit Sub
    enabled = SelectionIs(SEL_TEXT)
End Sub

Sub EnableWhenMultipleShapesSelected(control As IRibbonControl, ByRef enabled)
If Not SettingContextualButtons Then enabled = True: Exit Sub
    enabled = SelectionIs(SEL_MULTIPLE_SHAPES)
End Sub

Sub EnableWhenShapeOrText(control As IRibbonControl, ByRef enabled)
If Not SettingContextualButtons Then enabled = True: Exit Sub
    enabled = SelectionIs(SEL_SHAPES Or SEL_TEXT)
End Sub

Sub EnableWhenExactlyOneShape(control As IRibbonControl, ByRef enabled)
If Not SettingContextualButtons Then enabled = True: Exit Sub
    enabled = SelectionIs(SEL_ONE_SHAPE)
End Sub

Public Sub EnableWhenInTable(control As IRibbonControl, ByRef enabled)
If Not SettingContextualButtons Then enabled = True: Exit Sub
    enabled = SelectionIs(SEL_IN_TABLE)
End Sub



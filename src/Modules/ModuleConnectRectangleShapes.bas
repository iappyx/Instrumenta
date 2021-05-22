Attribute VB_Name = "ModuleConnectRectangleShapes"
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

Sub ConnectRectangleShapes(ShapeDirection As String)
    
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
    
    If ActiveWindow.Selection.ShapeRange.Count = 2 Then
    
    Dim Left1, Right1, Top1, Bottom1, Left2, Right2, Top2, Bottom2 As Double
    
    
    Left1 = ActiveWindow.Selection.ShapeRange(1).Left
    Right1 = Left1 + ActiveWindow.Selection.ShapeRange(1).Width
    Top1 = ActiveWindow.Selection.ShapeRange(1).Top
    Bottom1 = Top1 + ActiveWindow.Selection.ShapeRange(1).Height
    
    Left2 = ActiveWindow.Selection.ShapeRange(2).Left
    Right2 = Left2 + ActiveWindow.Selection.ShapeRange(2).Width
    Top2 = ActiveWindow.Selection.ShapeRange(2).Top
    Bottom2 = Top2 + ActiveWindow.Selection.ShapeRange(2).Height
    
    Set myDocument = Application.ActiveWindow.Selection.SlideRange
    
    Select Case ShapeDirection
    
    Case "RightToLeft"
        With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Right1, Y1:=Top1)
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Bottom1
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Bottom2
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Top2
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Top1
            '.ConvertToShape
            .ConvertToShape.Line.Visible = msoFalse
        End With
        
    Case "LeftToRight"
        With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Right2, Y1:=Top2)
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Bottom2
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Bottom1
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Top1
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Top2
            '.ConvertToShape
            .ConvertToShape.Line.Visible = msoFalse
        End With
        
     Case "BottomToTop"
        With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Left1, Y1:=Bottom1)
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Bottom1
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Top2
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Top2
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Bottom1
            '.ConvertToShape
            .ConvertToShape.Line.Visible = msoFalse
        End With
        
     Case "TopToBottom"
        With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Left2, Y1:=Bottom2)
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Bottom2
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Top1
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Top1
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Bottom2
            '.ConvertToShape
            .ConvertToShape.Line.Visible = msoFalse
        End With
        
        
    End Select
    
    Else
    MsgBox "Select two shapes."
    End If
    
    Else
    MsgBox "Select two shapes."
    End If
    
    
End Sub

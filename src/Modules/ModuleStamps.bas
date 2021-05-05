Attribute VB_Name = "ModuleStamps"
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

Sub GenerateStamp(StampTitleText As String, StampColor As Long)
    
    Set myDocument = Application.ActiveWindow
    
    Dim Stamp       As Object
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim NumberOfStamps As Long
    
    NumberOfStamps = 0
    For shapeNumber = 1 To myDocument.Selection.SlideRange.Shapes.Count
        
        If InStr(1, myDocument.Selection.SlideRange.Shapes(shapeNumber).Name, "Stamp") = 1 Then
            NumberOfStamps = NumberOfStamps + 1
            
        End If
        
    Next
    
    Set StampBackground = myDocument.Selection.SlideRange.Shapes.AddShape(msoShapeRoundedRectangle, 100, 100, 94, 26)
    
    With StampBackground
        .Line.Visible = False
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Name = "StampBackground" + Str(RandomNumber)
    End With
    
    Set StampBackgroundInner = myDocument.Selection.SlideRange.Shapes.AddShape(msoShapeRoundedRectangle, 102, 102, 90, 22)
    
    With StampBackgroundInner
        .Line.Visible = False
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Name = "StampBackgroundInner" + Str(RandomNumber)
    End With
    
    Set StampText = myDocument.Selection.SlideRange.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 94, 26)
    
    Application.ActiveWindow.Selection.SlideRange.Shapes(1).TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    
    With StampText
        .TextFrame.AutoSize = ppAutoSizeNone
        .TextFrame.HorizontalAnchor = msoAnchorCenter
        .TextFrame.VerticalAnchor = msoAnchorMiddle
        .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
        .TextFrame.TextRange = StampTitleText
        .TextFrame.MarginBottom = 0
        .TextFrame.MarginTop = 0
        .TextFrame.MarginLeft = 0
        .TextFrame.MarginRight = 0
        
        .TextFrame.TextRange.Font.Bold = msoTrue
        .TextFrame.TextRange.Font.Name = "Arial"
        .TextFrame.TextRange.Font.Size = 10
        .Line.Visible = False
        .Name = "StampText" + Str(RandomNumber)
    End With
    
    ActiveWindow.Selection.SlideRange(1).Shapes.Range(Array("StampBackground" + Str(RandomNumber), "StampBackgroundInner" + Str(RandomNumber), "StampText" + Str(RandomNumber))).Select
    CommandBars.ExecuteMso ("ShapesCombine")
    
    For Each Shape In ActiveWindow.Selection.ShapeRange
        
        Shape.Name = "Stamp" + Str(RandomNumber)
        Shape.Top = 5
        Shape.Left = Application.ActivePresentation.PageSetup.SlideWidth - (NumberOfStamps + 1) * (Shape.Width + 5)
        Shape.Fill.ForeColor.RGB = StampColor
        
    Next
    ActiveWindow.Selection.Unselect
    
End Sub

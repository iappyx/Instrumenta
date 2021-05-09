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

Sub MoveStampsOffSlide()
    Set myDocument = Application.ActiveWindow
    
    For shapeNumber = 1 To myDocument.Selection.SlideRange.Shapes.Count
        
        If InStr(1, myDocument.Selection.SlideRange.Shapes(shapeNumber).Name, "Stamp") = 1 Then
            
            myDocument.Selection.SlideRange.Shapes(shapeNumber).Tags.Add "OLDPOSITIONTOP", CStr(myDocument.Selection.SlideRange.Shapes(shapeNumber).Top)
            myDocument.Selection.SlideRange.Shapes(shapeNumber).Tags.Add "OLDPOSITIONLEFT", CStr(myDocument.Selection.SlideRange.Shapes(shapeNumber).Left)
            
            
            With myDocument.Selection.SlideRange.Shapes(shapeNumber)
            ShapeRight = (Application.ActivePresentation.PageSetup.SlideWidth - .Left - .Width)
            ShapeBottom = (Application.ActivePresentation.PageSetup.SlideHeight - .Top - .Height)
                             
            If .Left <= ShapeRight And .Left <= .Top And .Left <= ShapeBottom Then
            
            .Left = -5 - .Width
            
            ElseIf .Top <= ShapeRight And .Top <= ShapeBottom And .Top <= .Left Then
            
            .Top = -5 - .Height
            
            ElseIf ShapeRight <= ShapeBottom And ShapeRight <= .Left And ShapeRight <= .Top Then
            
            .Left = 5 + Application.ActivePresentation.PageSetup.SlideWidth
            
            Else
            
            .Top = 5 + Application.ActivePresentation.PageSetup.SlideHeight
            
            End If
            
            End With
            End If
        
    Next
    
End Sub

Sub MoveStampsOnSlide()
    Set myDocument = Application.ActiveWindow
    
    For shapeNumber = 1 To myDocument.Selection.SlideRange.Shapes.Count
        On Error Resume Next
        If InStr(1, myDocument.Selection.SlideRange.Shapes(shapeNumber).Name, "Stamp") = 1 Then
            myDocument.Selection.SlideRange.Shapes(shapeNumber).Top = CLng(myDocument.Selection.SlideRange.Shapes(shapeNumber).Tags("OLDPOSITIONTOP"))
            myDocument.Selection.SlideRange.Shapes(shapeNumber).Left = CLng(myDocument.Selection.SlideRange.Shapes(shapeNumber).Tags("OLDPOSITIONLEFT"))
            
        End If
        On Error GoTo 0
    Next
    
End Sub

Sub DeleteStampsOnSlide()
    Set myDocument = Application.ActiveWindow
    
    For shapeNumber = myDocument.Selection.SlideRange.Shapes.Count To 1 Step -1
        
        If InStr(1, myDocument.Selection.SlideRange.Shapes(shapeNumber).Name, "Stamp") = 1 Then
            myDocument.Selection.SlideRange.Shapes(shapeNumber).Delete
        End If
        
    Next
End Sub

Sub DeleteStampsOnAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For shapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            
            If InStr(1, PresentationSlide.Shapes(shapeNumber).Name, "Stamp") = 1 Then
                PresentationSlide.Shapes(shapeNumber).Delete
            End If
            
        Next
        
    Next
    
End Sub

Sub MoveStampsOnAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For shapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            On Error Resume Next
            If InStr(1, PresentationSlide.Shapes(shapeNumber).Name, "Stamp") = 1 Then
            PresentationSlide.Shapes(shapeNumber).Top = CLng(PresentationSlide.Shapes(shapeNumber).Tags("OLDPOSITIONTOP"))
            PresentationSlide.Shapes(shapeNumber).Left = CLng(PresentationSlide.Shapes(shapeNumber).Tags("OLDPOSITIONLEFT"))
            End If
            On Error GoTo 0
        Next
        
    Next
    
End Sub

Sub MoveStampsOffAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For shapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            
            If InStr(1, PresentationSlide.Shapes(shapeNumber).Name, "Stamp") = 1 Then
                
            PresentationSlide.Shapes(shapeNumber).Tags.Add "OLDPOSITIONTOP", CStr(PresentationSlide.Shapes(shapeNumber).Top)
            PresentationSlide.Shapes(shapeNumber).Tags.Add "OLDPOSITIONLEFT", CStr(PresentationSlide.Shapes(shapeNumber).Left)
            
            
            With PresentationSlide.Shapes(shapeNumber)
            ShapeRight = (Application.ActivePresentation.PageSetup.SlideWidth - .Left - .Width)
            ShapeBottom = (Application.ActivePresentation.PageSetup.SlideHeight - .Top - .Height)
                             
            If .Left <= ShapeRight And .Left <= .Top And .Left <= ShapeBottom Then
            
            .Left = -5 - .Width
            
            ElseIf .Top <= ShapeRight And .Top <= ShapeBottom And .Top <= .Left Then
            
            .Top = -5 - .Height
            
            ElseIf ShapeRight <= ShapeBottom And ShapeRight <= .Left And ShapeRight <= .Top Then
            
            .Left = 5 + Application.ActivePresentation.PageSetup.SlideWidth
            
            Else
            
            .Top = 5 + Application.ActivePresentation.PageSetup.SlideHeight
            
            End If
            
            End With
            
            End If
            
        Next
        
    Next
    
End Sub


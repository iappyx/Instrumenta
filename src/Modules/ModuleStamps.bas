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
    
    Set MyDocument = Application.ActiveWindow
    
    Dim Stamp       As Object
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim NumberOfStamps As Long
    
    NumberOfStamps = 0
    For ShapeNumber = 1 To MyDocument.Selection.SlideRange.Shapes.Count
        
        If InStr(1, MyDocument.Selection.SlideRange.Shapes(ShapeNumber).Name, "Stamp") = 1 Then
            NumberOfStamps = NumberOfStamps + 1
            
        End If
        
    Next
    
    Set StampBackground = MyDocument.Selection.SlideRange.Shapes.AddShape(msoShapeRoundedRectangle, 100, 100, 94, 26)
    
    With StampBackground
        .Line.visible = False
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Name = "StampBackground" + Str(RandomNumber)
    End With
    
    Set StampBackgroundInner = MyDocument.Selection.SlideRange.Shapes.AddShape(msoShapeRoundedRectangle, 102, 102, 90, 22)
    
    With StampBackgroundInner
        .Line.visible = False
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Name = "StampBackgroundInner" + Str(RandomNumber)
    End With
    
    Set StampText = MyDocument.Selection.SlideRange.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 94, 26)
    
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
        .Line.visible = False
        .Name = "StampText" + Str(RandomNumber)
    End With
    
    ActiveWindow.Selection.SlideRange(1).Shapes.Range(Array("StampBackground" + Str(RandomNumber), "StampBackgroundInner" + Str(RandomNumber), "StampText" + Str(RandomNumber))).Select
    CommandBars.ExecuteMso ("ShapesCombine")
    
    For Each shape In ActiveWindow.Selection.ShapeRange
        
        shape.Name = "Stamp" + Str(RandomNumber)
        shape.Top = 5
        shape.left = Application.ActivePresentation.PageSetup.SlideWidth - (NumberOfStamps + 1) * (shape.Width + 5)
        shape.Fill.ForeColor.RGB = StampColor
        shape.Tags.Add "INSTRUMENTA STAMP", StampTitleText
        
    Next
    ActiveWindow.Selection.Unselect
    
End Sub

Sub MoveStampsOffSlide()
    Set MyDocument = Application.ActiveWindow
    
    For ShapeNumber = 1 To MyDocument.Selection.SlideRange.Shapes.Count
        
        If InStr(1, MyDocument.Selection.SlideRange.Shapes(ShapeNumber).Name, "Stamp") = 1 Then
            
            MyDocument.Selection.SlideRange.Shapes(ShapeNumber).Tags.Add "INSTRUMENTA OLD POSITION TOP", CStr(MyDocument.Selection.SlideRange.Shapes(ShapeNumber).Top)
            MyDocument.Selection.SlideRange.Shapes(ShapeNumber).Tags.Add "INSTRUMENTA OLD POSITION LEFT", CStr(MyDocument.Selection.SlideRange.Shapes(ShapeNumber).left)
            
            
            With MyDocument.Selection.SlideRange.Shapes(ShapeNumber)
            ShapeRight = (Application.ActivePresentation.PageSetup.SlideWidth - .left - .Width)
            ShapeBottom = (Application.ActivePresentation.PageSetup.SlideHeight - .Top - .Height)
                             
            If .left <= ShapeRight And .left <= .Top And .left <= ShapeBottom Then
            
            .left = -5 - .Width
            
            ElseIf .Top <= ShapeRight And .Top <= ShapeBottom And .Top <= .left Then
            
            .Top = -5 - .Height
            
            ElseIf ShapeRight <= ShapeBottom And ShapeRight <= .left And ShapeRight <= .Top Then
            
            .left = 5 + Application.ActivePresentation.PageSetup.SlideWidth
            
            Else
            
            .Top = 5 + Application.ActivePresentation.PageSetup.SlideHeight
            
            End If
            
            End With
            End If
        
    Next
    
End Sub

Sub MoveStampsOnSlide()
    Set MyDocument = Application.ActiveWindow
    
    For ShapeNumber = 1 To MyDocument.Selection.SlideRange.Shapes.Count
        On Error Resume Next
        If InStr(1, MyDocument.Selection.SlideRange.Shapes(ShapeNumber).Name, "Stamp") = 1 Then
            MyDocument.Selection.SlideRange.Shapes(ShapeNumber).Top = CLng(MyDocument.Selection.SlideRange.Shapes(ShapeNumber).Tags("INSTRUMENTA OLD POSITION TOP"))
            MyDocument.Selection.SlideRange.Shapes(ShapeNumber).left = CLng(MyDocument.Selection.SlideRange.Shapes(ShapeNumber).Tags("INSTRUMENTA OLD POSITION LEFT"))
            
        End If
        On Error GoTo 0
    Next
    
End Sub

Sub DeleteStampsOnSlide()
    Set MyDocument = Application.ActiveWindow
    
    For ShapeNumber = MyDocument.Selection.SlideRange.Shapes.Count To 1 Step -1
        
        If InStr(1, MyDocument.Selection.SlideRange.Shapes(ShapeNumber).Name, "Stamp") = 1 Then
            MyDocument.Selection.SlideRange.Shapes(ShapeNumber).Delete
        End If
        
    Next
End Sub

Sub DeleteStampsOnAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For ShapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            
            If InStr(1, PresentationSlide.Shapes(ShapeNumber).Name, "Stamp") = 1 Then
                PresentationSlide.Shapes(ShapeNumber).Delete
            End If
            
        Next
        
    Next
    
End Sub

Sub MoveStampsOnAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For ShapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            On Error Resume Next
            If InStr(1, PresentationSlide.Shapes(ShapeNumber).Name, "Stamp") = 1 Then
            PresentationSlide.Shapes(ShapeNumber).Top = CLng(PresentationSlide.Shapes(ShapeNumber).Tags("INSTRUMENTA OLD POSITION TOP"))
            PresentationSlide.Shapes(ShapeNumber).left = CLng(PresentationSlide.Shapes(ShapeNumber).Tags("INSTRUMENTA OLD POSITION LEFT"))
            End If
            On Error GoTo 0
        Next
        
    Next
    
End Sub

Sub MoveStampsOffAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For ShapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            
            If InStr(1, PresentationSlide.Shapes(ShapeNumber).Name, "Stamp") = 1 Then
                
            PresentationSlide.Shapes(ShapeNumber).Tags.Add "INSTRUMENTA OLD POSITION TOP", CStr(PresentationSlide.Shapes(ShapeNumber).Top)
            PresentationSlide.Shapes(ShapeNumber).Tags.Add "INSTRUMENTA OLD POSITION LEFT", CStr(PresentationSlide.Shapes(ShapeNumber).left)
            
            
            With PresentationSlide.Shapes(ShapeNumber)
            ShapeRight = (Application.ActivePresentation.PageSetup.SlideWidth - .left - .Width)
            ShapeBottom = (Application.ActivePresentation.PageSetup.SlideHeight - .Top - .Height)
                             
            If .left <= ShapeRight And .left <= .Top And .left <= ShapeBottom Then
            
            .left = -5 - .Width
            
            ElseIf .Top <= ShapeRight And .Top <= ShapeBottom And .Top <= .left Then
            
            .Top = -5 - .Height
            
            ElseIf ShapeRight <= ShapeBottom And ShapeRight <= .left And ShapeRight <= .Top Then
            
            .left = 5 + Application.ActivePresentation.PageSetup.SlideWidth
            
            Else
            
            .Top = 5 + Application.ActivePresentation.PageSetup.SlideHeight
            
            End If
            
            End With
            
            End If
            
        Next
        
    Next
    
End Sub


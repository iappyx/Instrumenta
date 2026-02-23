Attribute VB_Name = "ModuleStepsCounter"
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

Sub GenerateStepsCounter()
    
    Set MyDocument = Application.ActiveWindow
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim NumberOfSteps As Long
    
    NumberOfSteps = 0
    For ShapeNumber = 1 To MyDocument.Selection.SlideRange.shapes.count
        
        If InStr(1, MyDocument.Selection.SlideRange.shapes(ShapeNumber).name, "StepsCounter") = 1 Then
            On Error Resume Next
            If CLng(MyDocument.Selection.SlideRange.shapes(ShapeNumber).TextFrame.textRange.text) > NumberOfSteps Then
                NumberOfSteps = CLng(MyDocument.Selection.SlideRange.shapes(ShapeNumber).TextFrame.textRange.text)
                MyDocument.Selection.SlideRange.shapes(ShapeNumber).PickUp
                CounterHeight = MyDocument.Selection.SlideRange.shapes(ShapeNumber).height
                CounterWidth = MyDocument.Selection.SlideRange.shapes(ShapeNumber).width
                CounterShape = MyDocument.Selection.SlideRange.shapes(ShapeNumber).AutoShapeType
                
            End If
            On Error GoTo 0
        End If
        
    Next
    
    Set StepsCounter = MyDocument.Selection.SlideRange.shapes.AddShape(msoShapeOval, Application.ActivePresentation.PageSetup.slideWidth - (22 * (NumberOfSteps + 1)), 5, 20, 20)
    
    With StepsCounter
        .line.visible = False
        .Fill.ForeColor.RGB = RGB(0, 112, 192)
        .Fill.Transparency = 0.1
        .name = "StepsCounter" + Str(RandomNumber)
        .Tags.Add "INSTRUMENTA STEPSCOUNTER", (NumberOfSteps + 1)
        
        With .TextFrame
            .marginBottom = 0
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .VerticalAnchor = msoAnchorMiddle
            
            With .textRange
                .Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
                .text = CStr(NumberOfSteps + 1)
                With .Font
                    .Size = 10
                    .color.RGB = RGB(255, 255, 255)
                End With
            End With
            
        End With
    End With
    
    If NumberOfSteps > 0 Then
        StepsCounter.AutoShapeType = CounterShape
        StepsCounter.width = CounterWidth
        StepsCounter.height = CounterHeight
        StepsCounter.Apply
    End If
    
End Sub

Sub SelectAllStepsCounter()
    
    Set MyDocument = Application.ActiveWindow
    Dim SlideShape, SlideShapeToCheck As PowerPoint.shape
    Dim selectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 0
    
    For Each SlideShapeToCheck In MyDocument.View.Slide.shapes
        
        If InStr(SlideShapeToCheck.name, "StepsCounter") = 1 Then
            
            If ShapeCount = 0 Then
                SlideShapeToCheck.Select msoTrue
            Else
                SlideShapeToCheck.Select msoFalse
            End If
            ShapeCount = ShapeCount + 1
            
        End If
        
    Next SlideShapeToCheck
    
End Sub

Sub GenerateCrossSlideStepsCounter()
    
    Set MyDocument = Application.ActiveWindow
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim NumberOfSteps As Long
    
    NumberOfSteps = 0
    
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For ShapeNumber = PresentationSlide.shapes.count To 1 Step -1
            
            If InStr(PresentationSlide.shapes(ShapeNumber).name, "CrossSlideStepsCounter") = 1 Then
                
                On Error Resume Next
                If CLng(PresentationSlide.shapes(ShapeNumber).TextFrame.textRange.text) > NumberOfSteps Then
                    NumberOfSteps = CLng(PresentationSlide.shapes(ShapeNumber).TextFrame.textRange.text)
                    PresentationSlide.shapes(ShapeNumber).PickUp
                    CounterHeight = PresentationSlide.shapes(ShapeNumber).height
                    CounterWidth = PresentationSlide.shapes(ShapeNumber).width
                    CounterShape = PresentationSlide.shapes(ShapeNumber).AutoShapeType
                    
                End If
                On Error GoTo 0
                
            End If
            
        Next
        
    Next
    
    Set StepsCounter = MyDocument.Selection.SlideRange.shapes.AddShape(msoShapeOval, Application.ActivePresentation.PageSetup.slideWidth - (22 * (NumberOfSteps + 1)), 5, 20, 20)
    
    With StepsCounter
        .line.visible = False
        .Fill.ForeColor.RGB = RGB(112, 192, 0)
        .Fill.Transparency = 0.1
        .name = "CrossSlideStepsCounter" + Str(RandomNumber)
        .Tags.Add "INSTRUMENTA CROSSSLIDE STEPSCOUNTER", (NumberOfSteps + 1)
        
        With .TextFrame
            .marginBottom = 0
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .VerticalAnchor = msoAnchorMiddle
            
            With .textRange
                .Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
                .text = CStr(NumberOfSteps + 1)
                With .Font
                    .Size = 10
                    .color.RGB = RGB(255, 255, 255)
                End With
            End With
            
        End With
    End With
    
    If NumberOfSteps > 0 Then
        StepsCounter.AutoShapeType = CounterShape
        StepsCounter.width = CounterWidth
        StepsCounter.height = CounterHeight
        StepsCounter.Apply
    End If
    
End Sub

Sub SelectAllCrossSlideStepsCounter()
    
    Set MyDocument = Application.ActiveWindow
    Dim SlideShape, SlideShapeToCheck As PowerPoint.shape
    Dim selectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 0
    
    For Each SlideShapeToCheck In MyDocument.View.Slide.shapes
        
        If InStr(SlideShapeToCheck.name, "CrossSlideStepsCounter") = 1 Then
            
            If ShapeCount = 0 Then
                SlideShapeToCheck.Select msoTrue
            Else
                SlideShapeToCheck.Select msoFalse
            End If
            ShapeCount = ShapeCount + 1
            
        End If
        
    Next SlideShapeToCheck
    
End Sub

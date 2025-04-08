Attribute VB_Name = "ModuleSlidesConvertToPictures"
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

Sub ConvertAllSlidesToPictures()
    
    ProgressForm.Show
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
        
        PresentationSlide.Copy
        PresentationSlide.Shapes.Range.Delete
        
        #If Mac Then
            Set ImageShape = PresentationSlide.Shapes.Paste
        #Else
            Set ImageShape = PresentationSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile)
        #End If
        
        ImageShape.Top = 0
        ImageShape.left = 0
        ImageShape.Width = Application.ActivePresentation.PageSetup.SlideWidth
        ImageShape.Height = Application.ActivePresentation.PageSetup.SlideHeight
        
        ImageShape.Copy
        ImageShape.Delete
        
        #If Mac Then
            Set ImageShape2 = PresentationSlide.Shapes.Paste
        #Else
            Set ImageShape2 = PresentationSlide.Shapes.PasteSpecial(ppPasteJPG)
        #End If
        
    Next PresentationSlide
    
    ProgressForm.Hide
    Unload ProgressForm
    
End Sub

Sub ConvertSelectedSlidesToPictures()
    
    TotalSelectedSlides = ActiveWindow.Selection.SlideRange.Count
    ProgressForm.Show
    
    If TotalSelectedSlides > 0 Then
        SlideIndex = 0
        
        ProgressForm.Show
        
        For Each PresentationSlide In ActiveWindow.Selection.SlideRange
            SlideIndex = SlideIndex + 1
            SetProgress (SlideIndex / TotalSelectedSlides * 100)
            
            PresentationSlide.Copy
            PresentationSlide.Shapes.Range.Delete
            
            #If Mac Then
                Set ImageShape = PresentationSlide.Shapes.Paste
            #Else
                Set ImageShape = PresentationSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile)
            #End If
            
            ImageShape.Top = 0
            ImageShape.left = 0
            ImageShape.Width = Application.ActivePresentation.PageSetup.SlideWidth
            ImageShape.Height = Application.ActivePresentation.PageSetup.SlideHeight
            
            ImageShape.Copy
            ImageShape.Delete
            
            #If Mac Then
                Set ImageShape2 = PresentationSlide.Shapes.Paste
            #Else
                Set ImageShape2 = PresentationSlide.Shapes.PasteSpecial(ppPasteJPG)
            #End If
            
        Next PresentationSlide
        
        ProgressForm.Hide
        Unload ProgressForm
        
    End If
    
End Sub

Sub ConvertSlidesToPictures()
    
    Select Case CallToSlideScopesForm()
        
        Case "cancel"
            
        Case "selected"
            
            ConvertSelectedSlidesToPictures
            
        Case "all"
            
            ConvertAllSlidesToPictures
            
    End Select
    
End Sub

Sub InsertWatermarkAndConvertSlidesToPictures()
    Dim Watermark   As shape
    Const PI = 3.14159265358979
    
    Select Case CallToSlideScopesForm()
        
        Case "cancel"
            
        Case "selected"
            
            TotalSelectedSlides = ActiveWindow.Selection.SlideRange.Count
            
            If TotalSelectedSlides > 0 Then
                SlideIndex = 0
                
                WatermarkText = InputBox("Please input watermark text", "Watermark", "CONFIDENTIAL")
                PredefinedColor = RGB(204, 0, 0)
                WatermarkTextColor = ColorDialog(PredefinedColor)
                
                ProgressForm.Show
                
                For Each PresentationSlide In ActiveWindow.Selection.SlideRange
                    
                    SlideIndex = SlideIndex + 1
                    SetProgress (SlideIndex / TotalSelectedSlides * 100)
                    
                    With PresentationSlide
                        
                        Set Watermark = .Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, left:=0, Top:=0, Width:=400, Height:=100)
                        Watermark.Width = Sqr(Application.ActivePresentation.PageSetup.SlideWidth * Application.ActivePresentation.PageSetup.SlideWidth + Application.ActivePresentation.PageSetup.SlideHeight * Application.ActivePresentation.PageSetup.SlideHeight)
                        Watermark.TextFrame.TextRange.Text = WatermarkText
                        Watermark.TextFrame.TextRange.Font.Size = 100
                        Watermark.TextFrame.HorizontalAnchor = msoAnchorCenter
                        Watermark.rotation = -Atn(Application.ActivePresentation.PageSetup.SlideHeight / Application.ActivePresentation.PageSetup.SlideWidth) * 180 / PI
                        Watermark.left = (Application.ActivePresentation.PageSetup.SlideWidth - Watermark.Width) / 2
                        Watermark.Top = (Application.ActivePresentation.PageSetup.SlideHeight - Watermark.Height) / 2
                        
                        Watermark.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = WatermarkTextColor
                        Watermark.TextFrame2.TextRange.Characters.Font.Fill.Transparency = 0.9
                        
                    End With
                Next PresentationSlide
                
                ProgressForm.Hide
                Unload ProgressForm
                
                ConvertSelectedSlidesToPictures
                
            End If
            
        Case "all"
            
            WatermarkText = InputBox("Please input watermark text", "Watermark", "CONFIDENTIAL")
            PredefinedColor = RGB(204, 0, 0)
            WatermarkTextColor = ColorDialog(PredefinedColor)
            
            ProgressForm.Show
            
            For Each PresentationSlide In ActivePresentation.Slides
                With PresentationSlide
                    
                    SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
                    
                    Set Watermark = .Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, left:=0, Top:=0, Width:=400, Height:=100)
                    Watermark.Width = Sqr(Application.ActivePresentation.PageSetup.SlideWidth * Application.ActivePresentation.PageSetup.SlideWidth + Application.ActivePresentation.PageSetup.SlideHeight * Application.ActivePresentation.PageSetup.SlideHeight)
                    Watermark.TextFrame.TextRange.Text = WatermarkText
                    Watermark.TextFrame.TextRange.Font.Size = 100
                    Watermark.TextFrame.HorizontalAnchor = msoAnchorCenter
                    Watermark.rotation = -Atn(Application.ActivePresentation.PageSetup.SlideHeight / Application.ActivePresentation.PageSetup.SlideWidth) * 180 / PI
                    Watermark.left = (Application.ActivePresentation.PageSetup.SlideWidth - Watermark.Width) / 2
                    Watermark.Top = (Application.ActivePresentation.PageSetup.SlideHeight - Watermark.Height) / 2
                    
                    Watermark.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = WatermarkTextColor
                    Watermark.TextFrame2.TextRange.Characters.Font.Fill.Transparency = 0.9
                    
                End With
            Next PresentationSlide
            
            ProgressForm.Hide
            Unload ProgressForm
            
            ConvertAllSlidesToPictures
            
    End Select
    
End Sub

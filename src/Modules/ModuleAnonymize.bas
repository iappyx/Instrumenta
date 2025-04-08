Attribute VB_Name = "ModuleAnonymize"
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

Sub AnonymizeWithLoremIpsum()
    
    Select Case CallToSlideScopesForm()
        
        Case "cancel"
            
        Case "selected"
            
            TotalSelectedSlides = ActiveWindow.Selection.SlideRange.Count
            ProgressForm.Show
            
            If TotalSelectedSlides > 0 Then
                SlideIndex = 0
                
                For Each PresentationSlide In ActiveWindow.Selection.SlideRange
                    SlideIndex = SlideIndex + 1
                    SetProgress (SlideIndex / TotalSelectedSlides * 100)
                    
                    For Each SlideShape In PresentationSlide.Shapes
                    
                        AnonymizeShapeWithLoremIpsum SlideShape
                    
                    Next SlideShape
                    
                Next PresentationSlide
                
            End If
            
            ProgressForm.Hide
            Unload ProgressForm
            
        Case "all"
            
            ProgressForm.Show
            
            For Each PresentationSlide In ActivePresentation.Slides
                
                SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
                
                For Each SlideShape In PresentationSlide.Shapes
                    
                    AnonymizeShapeWithLoremIpsum SlideShape
                    
                Next SlideShape
                
            Next PresentationSlide
            
            ProgressForm.Hide
            Unload ProgressForm
            
    End Select
    
End Sub

Sub AnonymizeShapeWithLoremIpsum(SlideShape)
    
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            AnonymizeShapeWithLoremIpsum SlideShapeChild
        Next
        
    Else
        
        If SlideShape.HasTextFrame Then
            
            For Each Paragraph In SlideShape.TextFrame2.TextRange.Paragraphs
                If Paragraph.Length > 1 Then
                Paragraph.Text = GetLoremIpsum(Paragraph.Words.Count, Paragraph.Length)
                End If
            Next
            
        End If
        
        If SlideShape.HasTable Then
            For TableRow = 1 To SlideShape.Table.Rows.Count
                For TableColumn = 1 To SlideShape.Table.Columns.Count
                    
                    For Each Paragraph In SlideShape.Table.Cell(TableRow, TableColumn).shape.TextFrame2.TextRange.Paragraphs
                        If Paragraph.Length > 1 Then
                        Paragraph.Text = GetLoremIpsum(Paragraph.Words.Count, Paragraph.Length)
                        End If
                    Next
                    
                Next
            Next
        End If
        
        If SlideShape.HasSmartArt Then
            
            For SlideShapeSmartArtNode = 1 To SlideShape.SmartArt.AllNodes.Count
                
                For Each Paragraph In SlideShape.SmartArt.AllNodes(SlideShapeSmartArtNode).TextFrame2.TextRange.Paragraphs
                    If Paragraph.Length > 1 Then
                    Paragraph.Text = GetLoremIpsum(Paragraph.Words.Count, Paragraph.Length)
                    End If
                Next
                
            Next
            
        End If
        
    End If
    
End Sub

Public Function GetLoremIpsum(NumberOfWords As Long, MaxLength As Long) As String
    
    If (NumberOfWords <= 0) Then
        GetLoremIpsum = ""
        Exit Function
    End If
    
    Dim LoremIpsumWords() As String
    Dim LoremResult As String
    Dim WordCount   As Long
    
    LoremIpsum = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus ac finibus purus. Phasellus et ultricies erat. Nullam maximus risus est, a pulvinar lectus pulvinar ut. Integer dictum malesuada sapien ac vulputate. Nam leo mauris, tincidunt quis dictum vel, semper nec est. Sed et dignissim tortor. Phasellus bibendum elit posuere erat malesuada ornare a sed odio. Integer purus lectus, gravida ac porttitor in, volutpat dictum sem. Pellentesque fermentum ante euismod dolor pellentesque, vitae vestibulum odio sagittis. In et massa massa."
    LoremIpsum = LoremIpsum + "Mauris maximus sem eget semper sollicitudin. Nullam gravida eros non scelerisque cursus. Sed non sem iaculis diam lacinia fermentum id vitae neque. Nulla facilisi. Vestibulum interdum ex non lorem tristique condimentum. Vestibulum facilisis tincidunt nulla at commodo. Ut pretium rhoncus lacus eget porttitor. Etiam quis euismod risus. Maecenas vel porta ante. Curabitur at rutrum eros, et vehicula ligula. Duis in maximus ante. Duis sed est in diam finibus venenatis."
    LoremIpsum = LoremIpsum + "Morbi sollicitudin felis sed scelerisque congue. Nullam vitae urna facilisis, consectetur urna non, ultricies mauris. Vivamus leo tortor, cursus vitae lacinia eget, varius et libero. Fusce luctus nec lectus sed dignissim. Donec malesuada ipsum in sagittis dictum. Nam vel augue id nulla porttitor consectetur. Duis nec enim id enim sagittis aliquam. Curabitur at nulla mi."
    LoremIpsum = LoremIpsum + "Praesent ac turpis eu elit auctor rhoncus. Mauris quis vehicula purus. Morbi sed neque leo. Sed ornare, ipsum et vulputate mattis, augue nisl feugiat magna, nec consequat elit risus eu est. Fusce viverra, urna vel porttitor vehicula, nulla nunc efficitur nunc, quis dapibus nulla ex quis ante. Nunc auctor iaculis sodales. Nunc vitae diam scelerisque, pretium ante vel, tincidunt velit. Sed nec congue arcu. Vestibulum vestibulum dolor sed nulla consequat vulputate. Donec nec dolor sed massa facilisis hendrerit. Curabitur dignissim vestibulum orci, sed facilisis neque condimentum id. Pellentesque erat nibh, euismod at dui quis, rutrum consectetur dolor."
    LoremIpsum = LoremIpsum + "Duis non ex nec lorem venenatis pellentesque. Ut euismod luctus tortor, sed consequat ipsum luctus sed. Duis at velit consectetur, commodo justo id, viverra tellus. Phasellus eu turpis non nisl porta suscipit et at ipsum. Mauris sodales purus vitae dolor hendrerit feugiat. Sed sit amet semper urna, a egestas ex. Phasellus mollis sodales augue at fermentum. Quisque aliquam scelerisque congue. In vitae hendrerit orci. Quisque ut luctus nisi. Donec sit amet mollis neque. Suspendisse vulputate tempus elit. Mauris quis turpis pellentesque, bibendum lectus eu, aliquam leo. Duis congue magna ac erat iaculis, eu bibendum orci finibus."
    LoremIpsum = LoremIpsum + "Ut volutpat maximus orci, vel ultrices turpis consequat in. Cras eu euismod odio, quis dapibus neque. Mauris ut dui id lacus tincidunt dapibus a eget lacus. Aenean imperdiet fringilla justo, in pellentesque sapien placerat a. Donec nisi augue, tempor eu blandit sed, efficitur et mi. Donec efficitur lectus non eros placerat, at egestas diam iaculis. Integer sodales turpis congue sagittis tempor. Donec nec orci sit amet augue sagittis gravida id vitae massa. Donec nec tincidunt velit. Integer nisl dolor, mollis ut ultrices quis, fermentum sed nisi. Ut aliquam nisi at orci ullamcorper, at malesuada orci sodales. Nunc ut molestie mauris. Donec rutrum aliquet velit, nec maximus urna tincidunt sed."
    LoremIpsum = LoremIpsum + "Donec rhoncus massa leo, sit amet tempus dui rutrum ac. Suspendisse at rutrum libero. Proin pharetra maximus mollis. Morbi molestie quis tortor sed consectetur. Aenean ullamcorper iaculis pharetra. Maecenas et blandit nisl, quis scelerisque nisl. Donec vel tempor sem, ac consequat justo. Pellentesque quis libero euismod, feugiat lacus et, finibus eros. Aenean finibus sit amet massa consectetur semper. Ut hendrerit euismod ipsum. Pellentesque lorem leo, vulputate non orci ut, convallis semper ex. Nunc fermentum tempor sagittis. Aliquam erat volutpat. Vivamus fringilla finibus ex sed pharetra. Quisque pharetra dictum lectus, sit amet dapibus eros accumsan eu. Pellentesque at lectus eu ipsum congue mollis."
    LoremIpsum = LoremIpsum + "Nunc ac condimentum justo. Phasellus vel massa aliquet, pulvinar ligula in, ornare enim. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Nulla molestie nisi nec posuere tincidunt. Cras eget bibendum ante, id facilisis augue. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Donec id turpis maximus, semper orci ac, tristique arcu. Sed euismod sapien sed nisl scelerisque suscipit. Pellentesque mollis volutpat orci quis eleifend. Curabitur et nisi est. Integer finibus commodo pretium."
    LoremIpsum = LoremIpsum + "Nunc dignissim tincidunt blandit. Sed quis arcu a lacus cursus mollis vitae nec eros. Ut dignissim cursus massa, nec elementum leo pellentesque ut. Aenean nec nunc scelerisque dui maximus consequat. Morbi diam augue, ullamcorper eget dictum id, venenatis vitae ipsum. Nulla facilisi. Aliquam mollis leo sed leo tempus aliquam. Donec a erat at justo rhoncus commodo ut eu erat. Ut vitae nisl rutrum, consectetur leo quis, laoreet diam. Sed metus leo, semper sit amet volutpat ut, placerat eu diam. Donec malesuada nunc ac pretium hendrerit."
    LoremIpsum = LoremIpsum + "Integer viverra pulvinar augue. Nulla et erat sed ante suscipit vulputate. Proin a iaculis nisl. Pellentesque convallis lorem sit amet euismod tincidunt. Pellentesque nisl mauris, dignissim sed imperdiet vel, tristique a orci. Integer ut scelerisque quam. Sed scelerisque lectus ut convallis malesuada. Morbi vehicula hendrerit magna in placerat."
    LoremIpsum = LoremIpsum + "Integer non interdum sapien. Praesent dictum risus erat, non iaculis dolor bibendum accumsan. Fusce fermentum ultricies ultrices. Ut condimentum elit vitae scelerisque euismod. Suspendisse massa ante, interdum in nisl quis, blandit."
    LoremIpsum = LoremIpsum + LoremIpsum + LoremIpsum + LoremIpsum + LoremIpsum
    
    LoremIpsumWords = Split(LoremIpsum, " ")
    
    If (NumberOfWords > UBound(LoremIpsumWords)) Then
        GetLoremIpsum = LoremIpsum
        Exit Function
    End If
    
    LoremResult = LoremIpsumWords(0)
    WordCount = 1
    
    Do While (WordCount < NumberOfWords)
    
        If (Len(LoremResult & " " & LoremIpsumWords(WordCount)) <= MaxLength) Or NumberOfWords <= 2 Then
            LoremResult = LoremResult & " " & LoremIpsumWords(WordCount)
        End If
        
        WordCount = WordCount + 1
    Loop
    
    GetLoremIpsum = LoremResult
    
End Function

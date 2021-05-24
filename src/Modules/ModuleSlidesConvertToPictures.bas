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


Sub ConvertSlidesToPictures()
    #If Mac Then
        MsgBox "This Function will not work on a Mac"
    #Else
        For Each PresentationSlide In ActivePresentation.Slides
            
            PresentationSlide.Copy
            PresentationSlide.Shapes.Range.Delete
            
            Set ImageShape = PresentationSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile)
            
            ImageShape.Top = 0
            ImageShape.Left = 0
            ImageShape.Width = Application.ActivePresentation.PageSetup.SlideWidth
            ImageShape.Height = Application.ActivePresentation.PageSetup.SlideHeight
            
            ImageShape.Copy
            ImageShape.Delete
            Set ImageShape2 = PresentationSlide.Shapes.PasteSpecial(ppPasteJPG)
            
        Next PresentationSlide
    
    #End If
End Sub

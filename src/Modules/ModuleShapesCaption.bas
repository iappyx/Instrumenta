Attribute VB_Name = "ModuleShapesCaption"
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

Public NumberOfTableCaptions, NumberOfShapeCaptions As Long

Sub InsertCaption()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
        
        NumberOfTableCaptions = 0
        NumberOfShapeCaptions = 0
        
        Dim GroupedCaption As Object
        RandomNumber = Round(Rnd() * 1000000, 0)
                
        ProgressForm.Show
        
        For Each PresentationSlide In ActivePresentation.Slides
            
            SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
            
            For Each SlideShape In PresentationSlide.Shapes
                
                CountCaptions SlideShape
                
            Next SlideShape
            
        Next PresentationSlide
        
        ProgressForm.Hide
        Unload ProgressForm
        
        Dim Caption As shape
        Dim CaptionNumber As shape
        
        For Each SlideShape In ActiveWindow.Selection.ShapeRange
            With SlideShape
                
                Set Caption = Application.ActiveWindow.View.Slide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, left:=0, Top:=0, Width:=400, Height:=100)
                Set CaptionNumber = Application.ActiveWindow.View.Slide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, left:=0, Top:=0, Width:=400, Height:=100)
                
                If .HasTable Then
                    CaptionNumber.Tags.Add "INSTRUMENTA TABLE CAPTION", Str(NumberOfTableCaptions)
                    CaptionNumberText = "Table " & Str(NumberOfTableCaptions + 1) & " - "
                Else
                    CaptionNumber.Tags.Add "INSTRUMENTA SHAPE CAPTION", Str(NumberOfShapeCaptions)
                    CaptionNumberText = "Figure" & Str(NumberOfShapeCaptions + 1) & " - "
                End If
                
                Caption.TextFrame.TextRange.Text = InputBox("Caption:", "Please enter caption")
                CaptionNumber.TextFrame.TextRange.Text = CaptionNumberText
                Caption.TextFrame.TextRange.Font.Size = 9
                CaptionNumber.TextFrame.TextRange.Font.Size = 9
                
                Caption.TextFrame.MarginBottom = 0
                Caption.TextFrame.MarginLeft = 0
                Caption.TextFrame.MarginRight = 0
                Caption.TextFrame.MarginTop = 0
                
                CaptionNumber.TextFrame.MarginBottom = 0
                CaptionNumber.TextFrame.MarginLeft = 0
                CaptionNumber.TextFrame.MarginRight = 0
                CaptionNumber.TextFrame.MarginTop = 0
                
                CaptionNumber.Width = 0
                CaptionNumber.TextFrame.WordWrap = msoFalse
                CaptionNumber.TextFrame.AutoSize = ppAutoSizeShapeToFitText
                Caption.Width = SlideShape.Width - CaptionNumber.Width
                CaptionNumber.left = SlideShape.left
                CaptionNumber.Top = SlideShape.Top + SlideShape.Height + 5
                Caption.left = SlideShape.left + CaptionNumber.Width
                Caption.Top = CaptionNumber.Top
                
                Caption.Name = "Caption" + Str(RandomNumber)
                CaptionNumber.Name = "CaptionNumber" + Str(RandomNumber)
                
                Set GroupedCaption = Application.ActiveWindow.View.Slide.Shapes.Range(Array("Caption" + Str(RandomNumber), "CaptionNumber" + Str(RandomNumber))).Group
                
            End With
        Next
        
    End If
    
End Sub

Sub CountCaptions(SlideShape)
    
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            CountCaptions SlideShapeChild
        Next
        
    Else
        
        If Not SlideShape.Tags("INSTRUMENTA TABLE CAPTION") = "" Then
            NumberOfTableCaptions = NumberOfTableCaptions + 1
        End If
        
        If Not SlideShape.Tags("INSTRUMENTA SHAPE CAPTION") = "" Then
            NumberOfShapeCaptions = NumberOfShapeCaptions + 1
        End If
        
    End If
    
End Sub

Sub ReNumberCaptions()
    
    NumberOfTableCaptions = 0
    NumberOfShapeCaptions = 0
    
    ProgressForm.Show
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
        
        For Each SlideShape In PresentationSlide.Shapes
            
            ReCountCaptions SlideShape
            
        Next SlideShape
        
    Next PresentationSlide
    
    ProgressForm.Hide
    Unload ProgressForm
    
End Sub

Sub ReCountCaptions(SlideShape)
    
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            ReCountCaptions SlideShapeChild
        Next
        
    Else
        
        If Not SlideShape.Tags("INSTRUMENTA TABLE CAPTION") = "" Then
            NumberOfTableCaptions = NumberOfTableCaptions + 1
            SlideShape.TextFrame.TextRange.Text = "Table " & NumberOfTableCaptions & " - "
        End If
        
        If Not SlideShape.Tags("INSTRUMENTA SHAPE CAPTION") = "" Then
            NumberOfShapeCaptions = NumberOfShapeCaptions + 1
            SlideShape.TextFrame.TextRange.Text = "Figure " & NumberOfShapeCaptions & " - "
        End If
        
    End If
    
End Sub

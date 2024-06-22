Attribute VB_Name = "ModuleObjectsLockAndUnlock"
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

Sub LockAspectRatioToggleSelectedShapes()
    
    Dim SlideShape As shape
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        
        If SlideShape.LockAspectRatio = msoTrue Then
            SlideShape.LockAspectRatio = msoFalse
        Else
            SlideShape.LockAspectRatio = msoTrue
        End If
        
    Next SlideShape
    
    
End Sub


Sub LockToggleSelectedShapes()
    
    #If Mac Then
    
    MsgBox "Locking or unlocking objects is not (yet) supported on Mac"
    
    #Else
    
    
    Dim SlideShape As shape
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        
        If SlideShape.Locked = msoTrue Then
            SlideShape.Locked = msoFalse
        Else
            SlideShape.Locked = msoTrue
        End If
        
    Next SlideShape
    
    #End If
    
End Sub


Sub LockToggleAllShapesOnAllSlides()
    
    #If Mac Then
    
    MsgBox "Locking or unlocking objects is not (yet) supported on Mac"
    
    #Else
    
    ProgressForm.Show
    
    For Each PresentationSlide In ActivePresentation.Slides
    
    SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
        
        For Each SlideShape In PresentationSlide.Shapes
            
            If SlideShape.Locked = msoTrue Then
                SlideShape.Locked = msoFalse
            Else
                SlideShape.Locked = msoTrue
            End If
            
        Next SlideShape
        
    Next PresentationSlide
    
    ProgressForm.Hide
    Unload ProgressForm
    
    #End If
    
End Sub

Sub LockAllShapesOnAllSlides()

    #If Mac Then
    
    MsgBox "Locking or unlocking objects is not (yet) supported on Mac"
    
    #Else
    
    ProgressForm.Show
    
    For Each PresentationSlide In ActivePresentation.Slides
    
    SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
        
        For Each SlideShape In PresentationSlide.Shapes

                SlideShape.Locked = msoTrue
            
        Next SlideShape
        
    Next PresentationSlide
    
    ProgressForm.Hide
    Unload ProgressForm
    
    #End If
    
End Sub

Sub UnLockAllShapesOnAllSlides()
    
    #If Mac Then
    
    MsgBox "Locking or unlocking objects is not (yet) supported on Mac"
    
    #Else
        
    ProgressForm.Show
    
    For Each PresentationSlide In ActivePresentation.Slides
    
    SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
        
        For Each SlideShape In PresentationSlide.Shapes

                SlideShape.Locked = msoFalse
            
        Next SlideShape
        
    Next PresentationSlide
    
    ProgressForm.Hide
    Unload ProgressForm
    
    #End If
    
End Sub

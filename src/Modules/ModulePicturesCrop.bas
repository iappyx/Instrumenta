Attribute VB_Name = "ModulePicturesCrop"
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

Sub ApplySameCropToSelectedImages()

    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No picture or shape selected."
        
    ElseIf MyDocument.Selection.ShapeRange.Count > 1 Then
        
        
        If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultTransformationMethod", "0") = 0 Then
        
        Set PictureShape = Application.ActiveWindow.Selection.ShapeRange(1)
        
        For i = 2 To Application.ActiveWindow.Selection.ShapeRange.Count
        
        Set PictureShape2 = Application.ActiveWindow.Selection.ShapeRange(i)
            
            Set TemporaryShape = PictureShape2.Duplicate
            
            TemporaryShape.ScaleHeight 1, msoTrue
            ScaledHeight = TemporaryShape.Height / PictureShape2.Height
            
            TemporaryShape.ScaleWidth 1, msoTrue
            ScaledWidth = TemporaryShape.Width / PictureShape2.Width
            
            TemporaryShape.Delete
        
        Select Case PictureShape2.Type
        Case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoLinkedPicture, msoPicture
        
        PictureShape2.PictureFormat.CropLeft = ScaledWidth * (PictureShape2.PictureFormat.Crop.PictureWidth - PictureShape.Width) / 2
        PictureShape2.PictureFormat.CropRight = ScaledWidth * (PictureShape2.PictureFormat.Crop.PictureWidth - PictureShape.Width) / 2
        PictureShape2.PictureFormat.CropTop = ScaledHeight * (PictureShape2.PictureFormat.Crop.PictureHeight - PictureShape.Height) / 2
        PictureShape2.PictureFormat.CropBottom = ScaledHeight * (PictureShape2.PictureFormat.Crop.PictureHeight - PictureShape.Height) / 2
        End Select
    
        
        Next i
        
        Else
        
         Set PictureShape = Application.ActiveWindow.Selection.ShapeRange(Application.ActiveWindow.Selection.ShapeRange.Count)
        
        For i = 1 To Application.ActiveWindow.Selection.ShapeRange.Count - 1
        
        Set PictureShape2 = Application.ActiveWindow.Selection.ShapeRange(i)
            
            Set TemporaryShape = PictureShape2.Duplicate
            
            TemporaryShape.ScaleHeight 1, msoTrue
            ScaledHeight = TemporaryShape.Height / PictureShape2.Height
            
            TemporaryShape.ScaleWidth 1, msoTrue
            ScaledWidth = TemporaryShape.Width / PictureShape2.Width
            
            TemporaryShape.Delete
        
        Select Case PictureShape2.Type
        Case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoLinkedPicture, msoPicture
        
        PictureShape2.PictureFormat.CropLeft = ScaledWidth * (PictureShape2.PictureFormat.Crop.PictureWidth - PictureShape.Width) / 2
        PictureShape2.PictureFormat.CropRight = ScaledWidth * (PictureShape2.PictureFormat.Crop.PictureWidth - PictureShape.Width) / 2
        PictureShape2.PictureFormat.CropTop = ScaledHeight * (PictureShape2.PictureFormat.Crop.PictureHeight - PictureShape.Height) / 2
        PictureShape2.PictureFormat.CropBottom = ScaledHeight * (PictureShape2.PictureFormat.Crop.PictureHeight - PictureShape.Height) / 2
        End Select
    
        
        Next i
        
        
        
        End If
        

        
    Else
        
        MsgBox "Please select more than one picture."
        
    End If


End Sub

Sub PictureCropToSlide()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No picture or shape selected."
        
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        
        Set PictureShape = Application.ActiveWindow.Selection.ShapeRange(1)
        
        Select Case PictureShape.Type
        Case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoLinkedPicture, msoPicture
            
            Set TemporaryShape = PictureShape.Duplicate
            
            TemporaryShape.ScaleHeight 1, msoTrue
            ScaledHeight = TemporaryShape.Height / PictureShape.Height
            
            TemporaryShape.ScaleWidth 1, msoTrue
            ScaledWidth = TemporaryShape.Width / PictureShape.Width
            
            TemporaryShape.Delete
            
            With PictureShape
                
                .PictureFormat.CropLeft = 0
                .PictureFormat.CropTop = 0
                .PictureFormat.CropBottom = 0
                .PictureFormat.CropRight = 0
                
                If .left < 0 Then
                    .PictureFormat.CropLeft = 0 - (.left * ScaledWidth)
                End If
                
                If .Top < 0 Then
                    .PictureFormat.CropTop = 0 - (.Top * ScaledHeight)
                End If
                
                If (.left + .Width) > Application.ActivePresentation.PageSetup.SlideWidth Then
                    .PictureFormat.CropRight = (.left + .Width - Application.ActivePresentation.PageSetup.SlideWidth) * ScaledWidth
                End If
                
                If (.Top + .Height) > Application.ActivePresentation.PageSetup.SlideHeight Then
                    .PictureFormat.CropBottom = (.Top + .Height - Application.ActivePresentation.PageSetup.SlideHeight) * ScaledHeight
                End If
                
            End With
            
        Case msoAutoShape, msoFreeform
                
                Set CropArea = Application.ActiveWindow.Selection.SlideRange.Shapes.AddShape(msoShapeRectangle, 0, 0, Application.ActivePresentation.PageSetup.SlideWidth, Application.ActivePresentation.PageSetup.SlideHeight)
                CropArea.Select msoFalse
                CommandBars.ExecuteMso ("ShapesIntersect")
            
        Case Else
            
            MsgBox "Selected shape is not a picture or compatible shape."
            
        End Select
        
    Else
        
        MsgBox "Please select one picture or shape."
        
    End If
    
End Sub

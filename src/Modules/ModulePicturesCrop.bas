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

#If Win64 Then

    Private Type GdiplusStartupInput
        GdiplusVersion As Long
        DebugEventCallback As LongPtr
        SuppressBackgroundThread As Long
        SuppressExternalCodecs As Long
    End Type
    
    Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" ( _
        token As LongPtr, _
        inputbuf As GdiplusStartupInput, _
        Optional ByVal outputbuf As LongPtr = 0) As Long
    
    Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" ( _
        ByVal token As LongPtr)
    
    Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus" ( _
        ByVal filename As LongPtr, _
        image As LongPtr) As Long
    
    Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" ( _
        ByVal image As LongPtr) As Long
    
    Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" ( _
        ByVal image As LongPtr, _
        width As Long) As Long
    
    Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" ( _
        ByVal image As LongPtr, _
        height As Long) As Long
    
    Private Declare PtrSafe Function GdipBitmapGetPixel Lib "gdiplus" ( _
        ByVal bitmap As LongPtr, _
        ByVal x As Long, _
        ByVal y As Long, _
        color As Long) As Long
#Else
    Private Type GdiplusStartupInput
        GdiplusVersion As Long
        DebugEventCallback As Long
        SuppressBackgroundThread As Long
        SuppressExternalCodecs As Long
    End Type
    
    Private Declare Function GdiplusStartup Lib "gdiplus" ( _
        token As Long, _
        inputbuf As GdiplusStartupInput, _
        Optional ByVal outputbuf As Long = 0) As Long
    
    Private Declare Sub GdiplusShutdown Lib "gdiplus" ( _
        ByVal token As Long)
    
    Private Declare Function GdipLoadImageFromFile Lib "gdiplus" ( _
        ByVal filename As Long, _
        image As Long) As Long
    
    Private Declare Function GdipDisposeImage Lib "gdiplus" ( _
        ByVal image As Long) As Long
    
    Private Declare Function GdipGetImageWidth Lib "gdiplus" ( _
        ByVal image As Long, _
        width As Long) As Long
    
    Private Declare Function GdipGetImageHeight Lib "gdiplus" ( _
        ByVal image As Long, _
        height As Long) As Long
    
    Private Declare Function GdipBitmapGetPixel Lib "gdiplus" ( _
        ByVal bitmap As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        color As Long) As Long
#End If


Sub ApplySameCropToSelectedImages()

    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No picture or shape selected."
        
    ElseIf MyDocument.Selection.ShapeRange.count > 1 Then
        
        
        If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultTransformationMethod", "0") = 0 Then
        
        Set PictureShape = Application.ActiveWindow.Selection.ShapeRange(1)
        
        For i = 2 To Application.ActiveWindow.Selection.ShapeRange.count
        
        Set PictureShape2 = Application.ActiveWindow.Selection.ShapeRange(i)
            
            Set TemporaryShape = PictureShape2.Duplicate
            
            TemporaryShape.ScaleHeight 1, msoTrue
            ScaledHeight = TemporaryShape.height / PictureShape2.height
            
            TemporaryShape.ScaleWidth 1, msoTrue
            ScaledWidth = TemporaryShape.width / PictureShape2.width
            
            TemporaryShape.Delete
        
        Select Case PictureShape2.Type
        Case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoLinkedPicture, msoPicture
        
        PictureShape2.PictureFormat.CropLeft = ScaledWidth * (PictureShape2.PictureFormat.Crop.PictureWidth - PictureShape.width) / 2
        PictureShape2.PictureFormat.CropRight = ScaledWidth * (PictureShape2.PictureFormat.Crop.PictureWidth - PictureShape.width) / 2
        PictureShape2.PictureFormat.CropTop = ScaledHeight * (PictureShape2.PictureFormat.Crop.PictureHeight - PictureShape.height) / 2
        PictureShape2.PictureFormat.CropBottom = ScaledHeight * (PictureShape2.PictureFormat.Crop.PictureHeight - PictureShape.height) / 2
        End Select
    
        
        Next i
        
        Else
        
         Set PictureShape = Application.ActiveWindow.Selection.ShapeRange(Application.ActiveWindow.Selection.ShapeRange.count)
        
        For i = 1 To Application.ActiveWindow.Selection.ShapeRange.count - 1
        
        Set PictureShape2 = Application.ActiveWindow.Selection.ShapeRange(i)
            
            Set TemporaryShape = PictureShape2.Duplicate
            
            TemporaryShape.ScaleHeight 1, msoTrue
            ScaledHeight = TemporaryShape.height / PictureShape2.height
            
            TemporaryShape.ScaleWidth 1, msoTrue
            ScaledWidth = TemporaryShape.width / PictureShape2.width
            
            TemporaryShape.Delete
        
        Select Case PictureShape2.Type
        Case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoLinkedPicture, msoPicture
        
        PictureShape2.PictureFormat.CropLeft = ScaledWidth * (PictureShape2.PictureFormat.Crop.PictureWidth - PictureShape.width) / 2
        PictureShape2.PictureFormat.CropRight = ScaledWidth * (PictureShape2.PictureFormat.Crop.PictureWidth - PictureShape.width) / 2
        PictureShape2.PictureFormat.CropTop = ScaledHeight * (PictureShape2.PictureFormat.Crop.PictureHeight - PictureShape.height) / 2
        PictureShape2.PictureFormat.CropBottom = ScaledHeight * (PictureShape2.PictureFormat.Crop.PictureHeight - PictureShape.height) / 2
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
        
    ElseIf MyDocument.Selection.ShapeRange.count = 1 Then
        
        Set PictureShape = Application.ActiveWindow.Selection.ShapeRange(1)
        
        Select Case PictureShape.Type
        Case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoLinkedPicture, msoPicture
            
            Set TemporaryShape = PictureShape.Duplicate
            
            TemporaryShape.ScaleHeight 1, msoTrue
            ScaledHeight = TemporaryShape.height / PictureShape.height
            
            TemporaryShape.ScaleWidth 1, msoTrue
            ScaledWidth = TemporaryShape.width / PictureShape.width
            
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
                
                If (.left + .width) > Application.ActivePresentation.PageSetup.SlideWidth Then
                    .PictureFormat.CropRight = (.left + .width - Application.ActivePresentation.PageSetup.SlideWidth) * ScaledWidth
                End If
                
                If (.Top + .height) > Application.ActivePresentation.PageSetup.SlideHeight Then
                    .PictureFormat.CropBottom = (.Top + .height - Application.ActivePresentation.PageSetup.SlideHeight) * ScaledHeight
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

Public Sub CropSelectedImageByDominantEdgeColor()

    #If Mac Then
        MsgBox "Crop by edge color is not (yet) supported on Mac"
        Exit Sub
    #End If

    If ActiveWindow Is Nothing Then Exit Sub
    If ActiveWindow.Selection Is Nothing Then Exit Sub
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select a picture shape first.", vbExclamation
        Exit Sub
    End If

    Dim shp As shape
    Set shp = ActiveWindow.Selection.ShapeRange(1)

    If shp.Type <> msoPicture And shp.Type <> msoLinkedPicture Then
        MsgBox "The selected object is not a picture.", vbExclamation
        Exit Sub
    End If

    CropImageByDominantEdgeColor shp

End Sub

Public Sub CropImageByDominantEdgeColor(shp As shape)
    #If Mac Then
        MsgBox "Crop by edge color is not (yet) supported on Mac"
        Exit Sub
    #End If
    
    With shp.PictureFormat
        .CropLeft = 0
        .CropRight = 0
        .CropTop = 0
        .CropBottom = 0
    End With
    
    Dim TempPath As String, tempFile As String
    TempPath = Environ$("TEMP") & "\"
    tempFile = TempPath & "instrumenta_crop_temp.png"
    
    Dim testCropAmount As Double
    Dim picWidth As Double
    
    picWidth = shp.PictureFormat.Crop.PictureWidth
    testCropAmount = picWidth * 0.1
    
    If testCropAmount < 10 Then testCropAmount = 10
    If testCropAmount > 99 Then testCropAmount = 100
    If testCropAmount > picWidth / 2 Then testCropAmount = picWidth / 2
    
    shp.Export tempFile, ppShapeFormatPNG
    Dim w1 As Long
    w1 = GetImageWidth(tempFile)
    
    shp.PictureFormat.CropLeft = testCropAmount
    shp.Export tempFile & "2.png", ppShapeFormatPNG
    Dim w2 As Long
    w2 = GetImageWidth(tempFile & "2.png")
    
    Dim pixelsPerPoint As Double
    pixelsPerPoint = (w1 - w2) / testCropAmount
    
    shp.PictureFormat.CropLeft = 0
    
    On Error Resume Next
    Kill tempFile & "2.png"
    On Error GoTo 0
    
    Dim token As LongPtr, inputBmp As LongPtr, si As GdiplusStartupInput
    si.GdiplusVersion = 1
    GdiplusStartup token, si, ByVal 0&
    
    If GdipLoadImageFromFile(StrPtr(tempFile), inputBmp) <> 0 Then
        MsgBox "Could not load exported image.", vbCritical
        GdiplusShutdown token
        Exit Sub
    End If
    
    Dim w As Long, h As Long
    GdipGetImageWidth inputBmp, w
    GdipGetImageHeight inputBmp, h
    
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim x As Long, y As Long, c As Long
    

    For x = 0 To w - 1
        c = GetPixelColor(inputBmp, x, 0)
        If Not dict.Exists(c) Then dict(c) = 0
        dict(c) = dict(c) + 1
        
        c = GetPixelColor(inputBmp, x, h - 1)
        If Not dict.Exists(c) Then dict(c) = 0
        dict(c) = dict(c) + 1
    Next x
    

    For y = 0 To h - 1
        c = GetPixelColor(inputBmp, 0, y)
        If Not dict.Exists(c) Then dict(c) = 0
        dict(c) = dict(c) + 1
        
        c = GetPixelColor(inputBmp, w - 1, y)
        If Not dict.Exists(c) Then dict(c) = 0
        dict(c) = dict(c) + 1
    Next y
    
    Dim bgColor As Long, maxCount As Long
    Dim key As Variant
    
    For Each key In dict.Keys
        If dict(key) > maxCount Then
            maxCount = dict(key)
            bgColor = CLng(key)
        End If
    Next key
    

    Dim trimTop As Long, trimBottom As Long
    Dim trimLeft As Long, trimRight As Long
    Dim found As Boolean
    
    For y = 0 To h - 1
        found = False
        For x = 0 To w - 1
            If GetPixelColor(inputBmp, x, y) <> bgColor Then
                found = True: Exit For
            End If
        Next x
        If found Then trimTop = y: Exit For
    Next y
    
    For y = h - 1 To 0 Step -1
        found = False
        For x = 0 To w - 1
            If GetPixelColor(inputBmp, x, y) <> bgColor Then
                found = True: Exit For
            End If
        Next x
        If found Then trimBottom = h - 1 - y: Exit For
    Next y
    
    For x = 0 To w - 1
        found = False
        For y = 0 To h - 1
            If GetPixelColor(inputBmp, x, y) <> bgColor Then
                found = True: Exit For
            End If
        Next y
        If found Then trimLeft = x: Exit For
    Next x
    
    For x = w - 1 To 0 Step -1
        found = False
        For y = 0 To h - 1
            If GetPixelColor(inputBmp, x, y) <> bgColor Then
                found = True: Exit For
            End If
        Next y
        If found Then trimRight = w - 1 - x: Exit For
    Next x
    
    GdipDisposeImage inputBmp
    GdiplusShutdown token
    
    With shp.PictureFormat
        .CropLeft = trimLeft / pixelsPerPoint
        .CropRight = trimRight / pixelsPerPoint
        .CropTop = trimTop / pixelsPerPoint
        .CropBottom = trimBottom / pixelsPerPoint
    End With
    
    On Error Resume Next
    Kill tempFile
    On Error GoTo 0
    
End Sub


Private Function GetImageWidth(filePath As String) As Long
    Dim token As LongPtr, bmp As LongPtr, si As GdiplusStartupInput, w As Long
    si.GdiplusVersion = 1
    GdiplusStartup token, si, ByVal 0&
    GdipLoadImageFromFile StrPtr(filePath), bmp
    GdipGetImageWidth bmp, w
    GdipDisposeImage bmp
    GdiplusShutdown token
    GetImageWidth = w
End Function


Private Function GetPixelColor(bmp As LongPtr, x As Long, y As Long) As Long
    Dim color As Long
    GdipBitmapGetPixel bmp, x, y, color
    GetPixelColor = color And &HFFFFFF
End Function

Attribute VB_Name = "ModuleEyedropper"
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

'Code contribution by FabPei (https://github.com/FabPei)


#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#Else
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#End If

Private Type POINTAPI
    x As Long
    y As Long
End Type


#If VBA7 And Win64 Then
    Private Function GetPixelColor() As Long
        Dim pt As POINTAPI
        Dim hdc As LongPtr
        Dim lColorRGB As Long
        
        GetCursorPos pt
        hdc = GetDC(0)
        lColorRGB = GetPixel(hdc, pt.x, pt.y)
        ReleaseDC 0, hdc
        
        GetPixelColor = lColorRGB
    End Function
#Else
    Private Function GetPixelColor() As Long
        Dim pt As POINTAPI
        Dim hdc As Long
        Dim lColorRGB As Long
        
        GetCursorPos pt
        hdc = GetDC(0)
        lColorRGB = GetPixel(hdc, pt.x, pt.y)
        ReleaseDC 0, hdc
        
        GetPixelColor = lColorRGB
    End Function
#End If

Public Sub ApplyPixelColorToFill()
    Dim oSh As shape
    Dim lColor As Long
    
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select one or more shapes.", vbExclamation
        Exit Sub
    End If
    
    lColor = GetPixelColor()
    
    For Each oSh In ActiveWindow.Selection.ShapeRange
        On Error Resume Next
        oSh.Fill.ForeColor.RGB = lColor
        On Error GoTo 0
    Next oSh
End Sub

Public Sub ApplyPixelColorToOutline()
    Dim oSh As shape
    Dim lColor As Long
    
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select one or more shapes.", vbExclamation
        Exit Sub
    End If
    
    lColor = GetPixelColor()
    
    For Each oSh In ActiveWindow.Selection.ShapeRange
        On Error Resume Next
        
        oSh.Line.visible = msoTrue
        oSh.Line.ForeColor.RGB = lColor
        
        On Error GoTo 0
    Next oSh
End Sub

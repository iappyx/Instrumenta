Attribute VB_Name = "ModuleObjects"
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

'Code contribution by o485 (https://github.com/o485):
'- GetRealTop
'- GetRealLeft
'- GetRealWidth
'- GetRealHeight
'- SetRealTop
'- SetRealLeft
'- SetRealWidth

Function GetRealTop(SlideShape As shape) As Single
    Dim rotation    As Double
    Dim radians     As Double
    Dim centerY     As Single
    
    rotation = SlideShape.rotation Mod 360
    radians = rotation * (3.14159265358979 / 180)
    centerY = SlideShape.Top + SlideShape.height / 2
    Select Case rotation
        Case 0, 180
            GetRealTop = SlideShape.Top
        Case 90, 270
            GetRealTop = centerY - SlideShape.width / 2
        Case Else
            GetRealTop = centerY - (SlideShape.height * Abs(Cos(radians)) + SlideShape.width * Abs(Sin(radians))) / 2
    End Select
End Function

Function GetRealLeft(SlideShape As shape) As Single
    Dim rotation    As Double
    Dim radians     As Double
    Dim centerX     As Single
    
    rotation = SlideShape.rotation Mod 360
    radians = rotation * (3.14159265358979 / 180)
    centerX = SlideShape.left + SlideShape.width / 2
    
    Select Case rotation
        Case 0, 180
            GetRealLeft = SlideShape.left
        Case 90, 270
            GetRealLeft = centerX - SlideShape.height / 2
        Case Else
            GetRealLeft = centerX - (SlideShape.width * Abs(Cos(radians)) + SlideShape.height * Abs(Sin(radians))) / 2
    End Select
End Function

Function GetRealWidth(SlideShape As shape) As Single
    Dim rotation    As Double
    Dim radians     As Double
    
    rotation = SlideShape.rotation Mod 360
    radians = rotation * (3.14159265358979 / 180)
    
    Select Case rotation
        Case 0, 180
            GetRealWidth = SlideShape.width
        Case 90, 270
            GetRealWidth = SlideShape.height
        Case Else
            GetRealWidth = SlideShape.width * Abs(Cos(radians)) + SlideShape.height * Abs(Sin(radians))
    End Select
End Function

Function GetRealHeight(SlideShape As shape) As Single
    Dim rotation    As Double
    Dim radians     As Double
    
    rotation = SlideShape.rotation Mod 360
    radians = rotation * (3.14159265358979 / 180)
    
    Select Case rotation
        Case 0, 180
            GetRealHeight = SlideShape.height
        Case 90, 270
            GetRealHeight = SlideShape.width
        Case Else
            GetRealHeight = SlideShape.height * Abs(Cos(radians)) + SlideShape.width * Abs(Sin(radians))
    End Select
End Function

Sub SetRealTop(SlideShape As shape, newRealTop As Single)
    Dim currentRealTop As Single
    Dim offset      As Single
    
    currentRealTop = GetRealTop(SlideShape)
    offset = newRealTop - currentRealTop
    
    SlideShape.Top = SlideShape.Top + offset
End Sub

Sub SetRealLeft(SlideShape As shape, newRealLeft As Single)
    Dim currentRealLeft As Single
    Dim offset      As Single
    
    currentRealLeft = GetRealLeft(SlideShape)
    offset = newRealLeft - currentRealLeft
    
    SlideShape.left = SlideShape.left + offset
End Sub

Sub SetRealWidth(SlideShape As shape, newRealWidth As Single)
    Dim rotation    As Double
    rotation = SlideShape.rotation Mod 360
    
    Select Case rotation
        Case 0
            SlideShape.width = newRealWidth
        Case 90
            SlideShape.left = SlideShape.left - (SlideShape.height - newRealWidth)
            SlideShape.height = newRealWidth
        Case 180
            SlideShape.left = SlideShape.left - (SlideShape.width - newRealWidth)
            SlideShape.width = newRealWidth
        Case 270
            SlideShape.height = newRealWidth
        Case Else
            Dim radians As Double
            Dim cosTheta As Double
            Dim sinTheta As Double
            
            radians = rotation * (3.14159265358979 / 180)
            cosTheta = Cos(radians)
            sinTheta = Sin(radians)
            
            aspectRatio = SlideShape.width / SlideShape.height
            
            SlideShape.width = newRealWidth / ((Abs(cosTheta) + (Abs(sinTheta)) / aspectRatio))
            SlideShape.height = SlideShape.width / aspectRatio
    End Select
End Sub

Sub SetRealHeight(SlideShape As shape, newRealHeight As Single)
    Dim rotation    As Double
    rotation = SlideShape.rotation Mod 360
    
    Select Case rotation
        Case 0
            SlideShape.height = newRealHeight
        Case 90
            SlideShape.width = newRealHeight
        Case 180
            SlideShape.Top = SlideShape.Top - (SlideShape.height - newRealHeight)
            SlideShape.height = newRealHeight
        Case 270
            SlideShape.Top = SlideShape.Top - (SlideShape.width - newRealHeight)
            SlideShape.width = newRealHeight
        Case Else
            Dim radians As Double
            Dim cosTheta As Double
            Dim sinTheta As Double
            
            radians = rotation * (3.14159265358979 / 180)
            cosTheta = Cos(radians)
            sinTheta = Sin(radians)
            
            aspectRatio = SlideShape.width / SlideShape.height
            
            SlideShape.height = newRealHeight / (Abs(cosTheta) + (Abs(sinTheta) * aspectRatio))
            SlideShape.width = SlideShape.height * aspectRatio
    End Select
End Sub

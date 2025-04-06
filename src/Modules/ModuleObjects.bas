Attribute VB_Name = "ModuleObjects"
'MIT License

'Copyright (c) 2021 iappyx
'Module contributed by o485 (https://github.com/o485)

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

Function GetRealTop(SlideShape As shape) As Single
    Dim rotation    As Double
    Dim radians     As Double
    Dim centerY     As Single
    
    rotation = SlideShape.rotation Mod 360
    radians = rotation * (3.14159265358979 / 180)
    centerY = SlideShape.Top + SlideShape.Height / 2
    
    Select Case rotation
        Case 0, 180
            GetRealTop = SlideShape.Top
        Case 90, 270
            GetRealTop = centerY - SlideShape.Width / 2
        Case Else
            GetRealTop = centerY - (SlideShape.Height * Abs(Cos(radians)) + SlideShape.Width * Abs(Sin(radians))) / 2
    End Select
End Function

Function GetRealLeft(SlideShape As shape) As Single
    Dim rotation    As Double
    Dim radians     As Double
    Dim centerX     As Single
    
    rotation = SlideShape.rotation Mod 360
    radians = rotation * (3.14159265358979 / 180)
    centerX = SlideShape.left + SlideShape.Width / 2
    
    Select Case rotation
        Case 0, 180
            GetRealLeft = SlideShape.left
        Case 90, 270
            GetRealLeft = centerX - SlideShape.Height / 2
        Case Else
            GetRealLeft = centerX - (SlideShape.Width * Abs(Cos(radians)) + SlideShape.Height * Abs(Sin(radians))) / 2
    End Select
End Function

Function GetRealWidth(SlideShape As shape) As Single
    Dim rotation    As Double
    Dim radians     As Double
    
    rotation = SlideShape.rotation Mod 360
    radians = rotation * (3.14159265358979 / 180)
    
    Select Case rotation
        Case 0, 180
            GetRealWidth = SlideShape.Width
        Case 90, 270
            GetRealWidth = SlideShape.Height
        Case Else
            GetRealWidth = SlideShape.Width * Abs(Cos(radians)) + SlideShape.Height * Abs(Sin(radians))
    End Select
End Function

Function GetRealHeight(SlideShape As shape) As Single
    Dim rotation    As Double
    Dim radians     As Double
    
    rotation = SlideShape.rotation Mod 360
    radians = rotation * (3.14159265358979 / 180)
    
    Select Case rotation
        Case 0, 180
            GetRealHeight = SlideShape.Height
        Case 90, 270
            GetRealHeight = SlideShape.Width
        Case Else
            GetRealHeight = SlideShape.Height * Abs(Cos(radians)) + SlideShape.Width * Abs(Sin(radians))
    End Select
End Function

Sub SetRealTop(SlideShape As shape, newRealTop As Single)
    Dim currentRealTop As Single
    Dim offset      As Single
    
    currentRealTop = GetRealTop(sh)
    offset = newRealTop - currentRealTop
    
    SlideShape.Top = SlideShape.Top + offset
End Sub

Sub SetRealLeft(SlideShape As shape, newRealLeft As Single)
    Dim currentRealLeft As Single
    Dim offset      As Single
    
    currentRealLeft = GetRealLeft(sh)
    offset = newRealLeft - currentRealLeft
    
    SlideShape.left = SlideShape.left + offset
End Sub

Sub SetRealWidth(SlideShape As shape, newRealWidth As Single)
    Dim rotation    As Double
    rotation = SlideShape.rotation Mod 360
    
    Select Case rotation
        Case 0
            SlideShape.Width = newRealWidth
        Case 90
            SlideShape.left = SlideShape.left - (SlideShape.Height - newRealWidth)
            SlideShape.Height = newRealWidth
        Case 180
            SlideShape.left = SlideShape.left - (SlideShape.Width - newRealWidth)
            SlideShape.Width = newRealWidth
        Case 270
            SlideShape.Height = newRealWidth
        Case Else
            Dim radians As Double
            Dim cosTheta As Double
            Dim sinTheta As Double
            
            radians = rotation * (3.14159265358979 / 180)
            cosTheta = Cos(radians)
            sinTheta = Sin(radians)
            
            aspectRatio = SlideShape.Width / SlideShape.Height
            
            SlideShape.Width = newRealWidth / ((Abs(cosTheta) + (Abs(sinTheta)) / aspectRatio))
            SlideShape.Height = SlideShape.Width / aspectRatio
    End Select
End Sub

Sub SetRealHeight(SlideShape As shape, newRealHeight As Single)
    Dim rotation    As Double
    rotation = SlideShape.rotation Mod 360
    
    Select Case rotation
        Case 0
            SlideShape.Height = newRealHeight
        Case 90
            SlideShape.Width = newRealHeight
        Case 180
            SlideShape.Top = SlideShape.Top - (SlideShape.Height - newRealHeight)
            SlideShape.Height = newRealHeight
        Case 270
            SlideShape.Top = SlideShape.Top - (SlideShape.Width - newRealHeight)
            SlideShape.Width = newRealHeight
        Case Else
            Dim radians As Double
            Dim cosTheta As Double
            Dim sinTheta As Double
            
            radians = rotation * (3.14159265358979 / 180)
            cosTheta = Cos(radians)
            sinTheta = Sin(radians)
            
            aspectRatio = SlideShape.Width / SlideShape.Height
            
            SlideShape.Height = newRealHeight / (Abs(cosTheta) + (Abs(sinTheta) * aspectRatio))
            SlideShape.Width = SlideShape.Height * aspectRatio
    End Select
End Sub

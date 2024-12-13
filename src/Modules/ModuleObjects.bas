Attribute VB_Name = "ModuleObjects"
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

Function GetRealTop(sh As Shape) As Single
    Dim rotation As Double
    Dim radians As Double
    Dim centerY As Single

    rotation = sh.rotation Mod 360
    radians = rotation * (3.14159265358979 / 180)
    centerY = sh.Top + sh.Height / 2

    Select Case rotation
        Case 0, 180
            GetRealTop = sh.Top
        Case 90, 270
            GetRealTop = centerY - sh.Width / 2
        Case Else
            GetRealTop = centerY - (sh.Height * Abs(Cos(radians)) + sh.Width * Abs(Sin(radians))) / 2
    End Select
End Function

Function GetRealLeft(sh As Shape) As Single
    Dim rotation As Double
    Dim radians As Double
    Dim centerX As Single
    
    rotation = sh.rotation Mod 360
    radians = rotation * (3.14159265358979 / 180)
    centerX = sh.left + sh.Width / 2

    Select Case rotation
        Case 0, 180
            GetRealLeft = sh.left
        Case 90, 270
            GetRealLeft = centerX - sh.Height / 2
        Case Else
            GetRealLeft = centerX - (sh.Width * Abs(Cos(radians)) + sh.Height * Abs(Sin(radians))) / 2
    End Select
End Function

Function GetRealWidth(sh As Shape) As Single
    Dim rotation As Double
    Dim radians As Double
    
    rotation = sh.rotation Mod 360
    radians = rotation * (3.14159265358979 / 180)

    Select Case rotation
        Case 0, 180
            GetRealWidth = sh.Width
        Case 90, 270
            GetRealWidth = sh.Height
        Case Else
            GetRealWidth = sh.Width * Abs(Cos(radians)) + sh.Height * Abs(Sin(radians))
    End Select
End Function

Function GetRealHeight(sh As Shape) As Single
    Dim rotation As Double
    Dim radians As Double
    
    rotation = sh.rotation Mod 360
    radians = rotation * (3.14159265358979 / 180)

    Select Case rotation
        Case 0, 180
            GetRealHeight = sh.Height
        Case 90, 270
            GetRealHeight = sh.Width
        Case Else
            GetRealHeight = sh.Height * Abs(Cos(radians)) + sh.Width * Abs(Sin(radians))
    End Select
End Function

Sub SetRealTop(sh As Shape, newRealTop As Single)
    Dim currentRealTop As Single
    Dim offset As Single

    currentRealTop = GetRealTop(sh)
    offset = newRealTop - currentRealTop

    sh.Top = sh.Top + offset
End Sub

Sub SetRealLeft(sh As Shape, newRealLeft As Single)
    Dim currentRealLeft As Single
    Dim offset As Single

    currentRealLeft = GetRealLeft(sh)
    offset = newRealLeft - currentRealLeft

    sh.left = sh.left + offset
End Sub

Sub SetRealWidth(sh As Shape, newRealWidth As Single)
    Dim rotation As Double
    rotation = sh.rotation Mod 360

    Select Case rotation
        Case 0
            sh.Width = newRealWidth
        Case 90
            sh.left = sh.left - (sh.Height - newRealWidth)
            sh.Height = newRealWidth
        Case 180
            sh.left = sh.left - (sh.Width - newRealWidth)
            sh.Width = newRealWidth
        Case 270
            sh.Height = newRealWidth
        Case Else
            Dim radians As Double
            Dim cosTheta As Double
            Dim sinTheta As Double

            radians = rotation * (3.14159265358979 / 180)
            cosTheta = Cos(radians)
            sinTheta = Sin(radians)

            aspectRatio = sh.Width / sh.Height
            
            sh.Width = newRealWidth / ((Abs(cosTheta) + (Abs(sinTheta)) / aspectRatio))
            sh.Height = sh.Width / aspectRatio
    End Select
End Sub

Sub SetRealHeight(sh As Shape, newRealHeight As Single)
    Dim rotation As Double
    rotation = sh.rotation Mod 360

    Select Case rotation
        Case 0
            sh.Height = newRealHeight
        Case 90
            sh.Width = newRealHeight
        Case 180
            sh.Top = sh.Top - (sh.Height - newRealHeight)
            sh.Height = newRealHeight
        Case 270
            sh.Top = sh.Top - (sh.Width - newRealHeight)
            sh.Width = newRealHeight
        Case Else
            Dim radians As Double
            Dim cosTheta As Double
            Dim sinTheta As Double

            radians = rotation * (3.14159265358979 / 180)
            cosTheta = Cos(radians)
            sinTheta = Sin(radians)

            aspectRatio = sh.Width / sh.Height
            
            sh.Height = newRealHeight / (Abs(cosTheta) + (Abs(sinTheta) * aspectRatio))
            sh.Width = sh.Height * aspectRatio
    End Select
End Sub

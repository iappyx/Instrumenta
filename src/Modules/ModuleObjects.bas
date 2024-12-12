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
    Dim centerY As Single

    rotation = sh.Rotation Mod 360
    centerY = sh.Top + sh.Height / 2

    Select Case rotation
        Case 0, 180
            GetRealTop = sh.Top
        Case 90, 270
            GetRealTop = centerY - sh.Width / 2
        Case Else
            GetRealTop = centerY - (sh.Height * Abs(Cos(rotation * (3.14159265358979 / 180))) + _
                                   sh.Width * Abs(Sin(rotation * (3.14159265358979 / 180)))) / 2
    End Select
End Function

Function GetRealLeft(sh As Shape) As Single
    Dim rotation As Double
    Dim centerX As Single

    rotation = sh.Rotation Mod 360
    centerX = sh.Left + sh.Width / 2

    Select Case rotation
        Case 0, 180
            GetRealLeft = sh.Left
        Case 90, 270
            GetRealLeft = centerX - sh.Height / 2
        Case Else
            GetRealLeft = centerX - (sh.Width * Abs(Cos(rotation * (3.14159265358979 / 180))) + _
                                     sh.Height * Abs(Sin(rotation * (3.14159265358979 / 180)))) / 2
    End Select
End Function

Function GetRealWidth(sh As Shape) As Single
    Dim rotation As Double
    rotation = sh.Rotation Mod 360

    Select Case rotation
        Case 0, 180
            GetRealWidth = sh.Width
        Case 90, 270
            GetRealWidth = sh.Height
        Case Else
            GetRealWidth = sh.Width * Abs(Cos(rotation * (3.14159265358979 / 180))) + _
                           sh.Height * Abs(Sin(rotation * (3.14159265358979 / 180)))
    End Select
End Function

Function GetRealHeight(sh As Shape) As Single
    Dim rotation As Double
    rotation = sh.Rotation Mod 360

    Select Case rotation
        Case 0, 180
            GetRealHeight = sh.Height
        Case 90, 270
            GetRealHeight = sh.Width
        Case Else
            GetRealHeight = sh.Height * Abs(Cos(rotation * (3.14159265358979 / 180))) + _
                            sh.Width * Abs(Sin(rotation * (3.14159265358979 / 180)))
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

    sh.Left = sh.Left + offset
End Sub

Sub SetRealWidth(sh As Shape, newRealWidth As Single)
    Dim rotation As Double
    rotation = sh.Rotation Mod 360

    Select Case rotation
        Case 0
            sh.Width = newRealWidth
        Case 90
            sh.Left = sh.Left - (sh.Height - newRealWidth)
            sh.Height = newRealWidth
        Case 180
            sh.Left = sh.Left - (sh.Width - newRealWidth)
            sh.Width = newRealWidth
        Case 270
            sh.Height = newRealWidth
        Case Else
            ' Dim cosTheta As Double, sinTheta As Double
            ' cosTheta = Abs(Cos(rotation * (3.14159265358979 / 180)))
            ' sinTheta = Abs(Sin(rotation * (3.14159265358979 / 180)))
            ' sh.Width = newRealWidth / (cosTheta + sinTheta)
    End Select
End Sub

Sub SetRealHeight(sh As Shape, newRealHeight As Single)
    Dim rotation As Double
    rotation = sh.Rotation Mod 360

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
            ' Dim cosTheta As Double, sinTheta As Double
            ' cosTheta = Abs(Cos(rotation * (3.14159265358979 / 180)))
            ' sinTheta = Abs(Sin(rotation * (3.14159265358979 / 180)))
            ' sh.Height = newRealHeight / (cosTheta + sinTheta)
    End Select
End Sub

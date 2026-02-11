VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColorManagerForm 
   Caption         =   "Color replacer"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10710
   OleObjectBlob   =   "ColorManagerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ColorManagerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Option Explicit
Private ColorData As String
Private TotalColors As Long
Private SelectedColorRGB As Long

Public Sub ShowColors(colors() As ColorInfo, colorCount As Long, elapsed As Double, scope As String)
       
    Dim i As Long
    Dim colorStr As String
    
    For i = 1 To colorCount
        colorStr = colorStr & colors(i).RGB & "|" & _
                              colors(i).RedValue & "|" & _
                              colors(i).GreenValue & "|" & _
                              colors(i).BlueValue & "|" & _
                              colors(i).HexValue & "|" & _
                              colors(i).UsageCount & "|" & _
                              colors(i).UsageTypes & "||"
    Next i
    
    ColorData = colorStr
    TotalColors = colorCount
    
    Me.Caption = "Color Manager - " & colorCount & " unique colors found in " & scope
    
    With lstColors
        .ColumnCount = 4
        .ColumnWidths = "80;80;60;150"
    End With
    
    PopulateColorList
    
    Me.Show vbModeless
End Sub



Private Sub lblNewColorPreview_Click()
    
    Dim r As Long, g As Long, b As Long
    Dim newColor As Long
    
    newColor = ColorDialog(RGB(0, 0, 0))
    
    r = newColor Mod 256
    g = (newColor \ 256) Mod 256
    b = (newColor \ 65536) Mod 256
    
    txtNewColor.Text = RGBToHex(r, g, b)
    lblNewColorPreview.BackColor = newColor
End Sub

Private Sub UserForm_Initialize()
    cmdReplaceColor.Enabled = False
    lblColorPreview.BackColor = vbWhite
    lblColorPreview.BorderStyle = fmBorderStyleSingle
    
    lblNewColorPreview.BackColor = vbWhite
    lblNewColorPreview.BorderStyle = fmBorderStyleSingle
End Sub

Private Sub PopulateColorList()
    Dim rows() As String
    Dim fields() As String
    Dim i As Long
    Dim rgbStr As String
    
    lstColors.Clear
    
    rows = Split(ColorData, "||")
    
    For i = 0 To UBound(rows)
        If Len(Trim(rows(i))) > 0 Then
            fields = Split(rows(i), "|")
            If UBound(fields) >= 6 Then

                lstColors.AddItem fields(4)  ' Hex
                

                rgbStr = fields(1) & "," & fields(2) & "," & fields(3)
                lstColors.List(lstColors.ListCount - 1, 1) = rgbStr
                

                lstColors.List(lstColors.ListCount - 1, 2) = fields(5)
                
                Dim usageStr As String
                usageStr = fields(6)
                If Len(usageStr) > 30 Then
                    usageStr = left(usageStr, 27) & "..."
                End If
                lstColors.List(lstColors.ListCount - 1, 3) = usageStr
            End If
        End If
    Next i
    
    If lstColors.ListCount > 0 Then
        lstColors.ListIndex = 0
    End If
End Sub

Private Sub lstColors_Click()
    If lstColors.ListIndex < 0 Then Exit Sub
    
    Dim rows() As String
    Dim fields() As String
    
    rows = Split(ColorData, "||")
    
    If lstColors.ListIndex <= UBound(rows) Then
        fields = Split(rows(lstColors.ListIndex), "|")
        
        If UBound(fields) >= 6 Then
            SelectedColorRGB = CLng(fields(0))
            
            lblColorPreview.BackColor = SelectedColorRGB
            lblColorInfo.Caption = "Selected Color:" & vbCrLf & _
                                  "Hex: " & fields(4) & vbCrLf & _
                                  "RGB: " & fields(1) & ", " & fields(2) & ", " & fields(3) & vbCrLf & _
                                  "Used " & fields(5) & " time(s) in: " & fields(6)
            
            txtOldColor.Text = fields(4)
            
            cmdReplaceColor.Enabled = True
        End If
    End If
End Sub

Private Sub cmdPickNewColor_Click()

    Dim r As Long, g As Long, b As Long
    Dim newColor As Long
    
    newColor = ColorDialog(RGB(0, 0, 0))
    
    r = newColor Mod 256
    g = (newColor \ 256) Mod 256
    b = (newColor \ 65536) Mod 256
    
    txtNewColor.Text = RGBToHex(r, g, b)
    lblNewColorPreview.BackColor = newColor
   
End Sub

Private Sub cmdReplaceColor_Click()
    If lstColors.ListIndex < 0 Then
        MsgBox "Please select a color to replace.", vbExclamation
        Exit Sub
    End If
    
    If Len(Trim(txtNewColor.Text)) = 0 Then
        MsgBox "Please enter a new color.", vbExclamation
        Exit Sub
    End If
    
    Dim newRGB As Long
    newRGB = ParseColorInput(txtNewColor.Text)
    
    If newRGB = -1 Then
        MsgBox "Invalid color format. Use hex (#FF0000) or RGB (255,0,0).", vbExclamation
        Exit Sub
    End If
    
    Dim msg As String
    msg = "Replace all instances of " & txtOldColor.Text & " with the new color in " & RecolorSlideScope & "?" & vbCrLf & vbCrLf
    msg = msg & "This action cannot be undone (except via Ctrl+Z)."
    
    If MsgBox(msg, vbQuestion + vbYesNo, "Confirm Color Replacement") = vbYes Then
        
        With ModuleColorScanner.RecolorUserPerm
        .AllowFill = chkFill.Value
        .AllowLine = chkLine.Value
        .AllowText = chkText.Value
        .AllowTableFill = chkTableFill.Value
        .AllowTableBorders = chkTableBorders.Value
        .AllowChart = chkChart.Value
        .AllowBackground = chkBackground.Value
        End With

               
        Me.Hide
        
        
        If RecolorSlideScope = "selected shapes" Then
        ModuleColorScanner.ReplaceColorInSelectedShapes SelectedColorRGB, newRGB
        Unload Me
        ModuleColorScanner.ScanColorsInSelectedShapes
        Else
        ModuleColorScanner.ReplaceColor SelectedColorRGB, newRGB, RecolorSlideScope
        Unload Me
        ModuleColorScanner.ScanAndManageColors
        End If
               

     
        
    End If
End Sub


Private Sub cmdClose_Click()
    RecolorSlideScope = ""
    Unload Me
End Sub

Private Function ParseColorInput(colorStr As String) As Long
    Dim r As Long, g As Long, b As Long
    Dim parts() As String
    
    On Error GoTo ParseError
    
    colorStr = Trim(colorStr)
    

    If left(colorStr, 1) = "#" Then

        colorStr = Mid(colorStr, 2)
        
        If Len(colorStr) = 6 Then
            r = CLng("&H" & Mid(colorStr, 1, 2))
            g = CLng("&H" & Mid(colorStr, 3, 2))
            b = CLng("&H" & Mid(colorStr, 5, 2))
            
            ParseColorInput = RGB(r, g, b)
            Exit Function
        End If
    End If
    
    If InStr(colorStr, ",") > 0 Then
        parts = Split(colorStr, ",")
        
        If UBound(parts) = 2 Then
            r = CLng(Trim(parts(0)))
            g = CLng(Trim(parts(1)))
            b = CLng(Trim(parts(2)))
            
            If r >= 0 And r <= 255 And g >= 0 And g <= 255 And b >= 0 And b <= 255 Then
                ParseColorInput = RGB(r, g, b)
                Exit Function
            End If
        End If
    End If
    
ParseError:
    ParseColorInput = -1
End Function

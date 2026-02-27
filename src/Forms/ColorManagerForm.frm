Attribute VB_Name = "ColorManagerForm"
Attribute VB_Base = "0{AF971A0D-B90A-44BD-A8D6-E3D32ED3DD80}{F284DD11-89BA-42B2-B8D9-FBC8D0C40AC9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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

Option Explicit
Private ColorData As String
Private TotalColors As Long
Private SelectedColorRGB As Long

Private Sub UpdateReplaceButtonState()
    cmdReplaceColor.enabled = (chkFill.value Or chkLine.value Or chkText.value Or chkTableFill.value Or chkTableBorders.value Or chkChart.value Or chkBackground.value)
End Sub


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
    
    Me.caption = "Color Manager - " & colorCount & " unique colors found in " & scope
    
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
    
    txtNewColor.text = RGBToHex(r, g, b)
    lblNewColorPreview.BackColor = newColor
End Sub

Private Sub UserForm_Initialize()
    cmdReplaceColor.enabled = False
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
    Dim usageStr As String
    
    rows = Split(ColorData, "||")
    
    If lstColors.ListIndex <= UBound(rows) Then
        fields = Split(rows(lstColors.ListIndex), "|")
        
        If UBound(fields) >= 6 Then
            SelectedColorRGB = CLng(fields(0))
            
            lblColorPreview.BackColor = SelectedColorRGB
            lblColorInfo.caption = "Selected Color:" & vbCrLf & _
                                  "Hex: " & fields(4) & vbCrLf & _
                                  "RGB: " & fields(1) & ", " & fields(2) & ", " & fields(3) & vbCrLf & _
                                  "Used " & fields(5) & " time(s) in: " & fields(6)
            
            txtOldColor.text = fields(4)
            cmdReplaceColor.enabled = True
            
            chkFill.value = False
            chkLine.value = False
            chkText.value = False
            chkTableFill.value = False
            chkTableBorders.value = False
            chkChart.value = False
            chkBackground.value = False
            
            usageStr = fields(6)
            
            chkFill.enabled = (InStr(1, usageStr, "Shape Fill", vbTextCompare) > 0) _
                           Or (InStr(1, usageStr, "Gradient Fill", vbTextCompare) > 0)
            
            chkLine.enabled = (InStr(1, usageStr, "Line/Border", vbTextCompare) > 0)
            
            chkText.enabled = (InStr(1, usageStr, "Font Color", vbTextCompare) > 0)
            
            chkTableFill.enabled = (InStr(1, usageStr, "Table Cell Fill", vbTextCompare) > 0)
            
            chkTableBorders.enabled = (InStr(1, usageStr, "Table Border", vbTextCompare) > 0)
            
            chkChart.enabled = (InStr(1, usageStr, "Chart Series", vbTextCompare) > 0) _
                            Or (InStr(1, usageStr, "Chart Point", vbTextCompare) > 0)
            
            chkBackground.enabled = (InStr(1, usageStr, "Slide Background", vbTextCompare) > 0)
            
            chkFill.value = chkFill.enabled
            chkLine.value = chkLine.enabled
            chkText.value = chkText.enabled
            chkTableFill.value = chkTableFill.enabled
            chkTableBorders.value = chkTableBorders.enabled
            chkChart.value = chkChart.enabled
            chkBackground.value = chkBackground.enabled
            
            chkFill.value = chkFill.enabled
            chkLine.value = chkLine.enabled
            chkText.value = chkText.enabled
            chkTableFill.value = chkTableFill.enabled
            chkTableBorders.value = chkTableBorders.enabled
            chkChart.value = chkChart.enabled
            chkBackground.value = chkBackground.enabled
            
            UpdateReplaceButtonState
            
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
    
    txtNewColor.text = RGBToHex(r, g, b)
    lblNewColorPreview.BackColor = newColor
   
End Sub

Private Sub cmdReplaceColor_Click()
    If lstColors.ListIndex < 0 Then
        MsgBox "Please select a color to replace.", vbExclamation
        Exit Sub
    End If
    
    If Len(Trim(txtNewColor.text)) = 0 Then
        MsgBox "Please enter a new color.", vbExclamation
        Exit Sub
    End If
    
    Dim newRGB As Long
    newRGB = ParseColorInput(txtNewColor.text)
    
    If newRGB = -1 Then
        MsgBox "Invalid color format. Use hex (#FF0000) or RGB (255,0,0).", vbExclamation
        Exit Sub
    End If
    
    Dim msg As String
    msg = "Replace all instances of " & txtOldColor.text & " with the new color in " & RecolorSlideScope & "?" & vbCrLf & vbCrLf
    msg = msg & "This action cannot be undone (except via Ctrl+Z)."
    
    If MsgBox(msg, vbQuestion + vbYesNo, "Confirm Color Replacement") = vbYes Then
        
        With ModuleColorScanner.RecolorUserPerm
        .AllowFill = chkFill.value
        .AllowLine = chkLine.value
        .AllowText = chkText.value
        .AllowTableFill = chkTableFill.value
        .AllowTableBorders = chkTableBorders.value
        .AllowChart = chkChart.value
        .AllowBackground = chkBackground.value
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

Private Sub chkFill_Click()
    UpdateReplaceButtonState
End Sub

Private Sub chkLine_Click()
    UpdateReplaceButtonState
End Sub

Private Sub chkText_Click()
    UpdateReplaceButtonState
End Sub

Private Sub chkTableFill_Click()
    UpdateReplaceButtonState
End Sub

Private Sub chkTableBorders_Click()
    UpdateReplaceButtonState
End Sub

Private Sub chkChart_Click()
    UpdateReplaceButtonState
End Sub

Private Sub chkBackground_Click()
    UpdateReplaceButtonState
End Sub


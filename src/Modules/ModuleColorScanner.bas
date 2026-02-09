Attribute VB_Name = "ModuleColorScanner"
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

Public Type ColorInfo
    RGB As Long
    RedValue As Long
    GreenValue As Long
    BlueValue As Long
    HexValue As String
    UsageCount As Long
    UsageTypes As String
End Type

Private colors() As ColorInfo
Private colorCount As Long

Public Sub ScanAndManageColors()
    Dim startTime As Double
    Dim SlideScope As String
    
    startTime = Timer
    
    SlideScope = CallToSlideScopesForm()
    
    Select Case SlideScope
        Case "cancel"
            Exit Sub
            
        Case "selected"
            ReDim colors(1 To 1000) As ColorInfo
            colorCount = 0

            ScanSelectedSlides
            ConsolidateColors
            ShowColorResults startTime, "selected slides"
            
        Case "all"
            ReDim colors(1 To 1000) As ColorInfo
            colorCount = 0
            
            ScanAllSlides
            ConsolidateColors
            ShowColorResults startTime, "all slides"
            
    End Select
End Sub

Private Sub ShowColorResults(startTime As Double, scope As String)
    Dim elapsed As Double
    elapsed = Timer - startTime
    
    If colorCount = 0 Then
        MsgBox "No colors found in " & scope & ".", vbInformation
    Else
        ColorManagerForm.ShowColors colors, colorCount, elapsed, scope
    End If
End Sub

Private Sub ScanSelectedSlides()
    Dim sld As Slide
    Dim shp As shape
    Dim slideIndex As Long
    
    On Error Resume Next
    
    For Each sld In ActiveWindow.Selection.SlideRange
        ScanSlideBackground sld
        
        For Each shp In sld.Shapes
            ScanShape shp
        Next shp
    Next sld
    
    On Error GoTo 0
End Sub

Private Sub ScanAllSlides()
    Dim sld As Slide
    Dim shp As shape
    
    On Error Resume Next
    
    For Each sld In ActivePresentation.Slides
        ScanSlideBackground sld
        For Each shp In sld.Shapes
            ScanShape shp
        Next shp
    Next sld
    
    On Error GoTo 0
End Sub

Private Sub ScanSlideBackground(sld As Slide)
    On Error Resume Next
    
If sld.Background.Fill.visible Then
    If sld.Background.Fill.ForeColor.Type = msoColorTypeRGB _
       Or sld.Background.Fill.ForeColor.Type = msoColorTypeScheme Then
        AddColor sld.Background.Fill.ForeColor.RGB, "Slide Background"
    End If
End If
    
    On Error GoTo 0
End Sub

Private Sub ScanShape(shp As shape)
    On Error Resume Next
    
    If shp.Type = msoGroup Then
        Dim subShape As shape
        For Each subShape In shp.GroupItems
            ScanShape subShape
        Next subShape
        Exit Sub
    End If
    
        If shp.Fill.visible Then
            Select Case shp.Fill.Type
                Case msoFillSolid
                    If shp.Fill.ForeColor.Type = msoColorTypeRGB _
                       Or shp.Fill.ForeColor.Type = msoColorTypeScheme Then
                        AddColor shp.Fill.ForeColor.RGB, "Shape Fill"
                    End If
            Case msoFillGradient
                    If shp.Fill.GradientStops.count > 0 Then
                        Dim i As Long
                        For i = 1 To shp.Fill.GradientStops.count
                            If shp.Fill.GradientStops(i).Color.Type = msoColorTypeRGB _
                               Or shp.Fill.GradientStops(i).Color.Type = msoColorTypeScheme Then
                                AddColor shp.Fill.GradientStops(i).Color.RGB, "Gradient Fill"
                            End If
                        Next i
End If
        End Select
    End If
    
    If shp.Line.visible Then
        If shp.Line.ForeColor.Type = msoColorTypeRGB _
           Or shp.Line.ForeColor.Type = msoColorTypeScheme Then
            AddColor shp.Line.ForeColor.RGB, "Line/Border"
        End If
    End If
    
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            ScanTextColors shp.TextFrame.textRange
        End If
    End If
    
    If shp.HasTable Then
        ScanTableColors shp.table
    End If
    
    If shp.HasChart Then
        ScanChartColors shp.Chart
    End If
    
    On Error GoTo 0
End Sub

Private Sub ScanTextColors(txtRange As textRange)
    On Error Resume Next
    
    Dim i As Long

    For i = 1 To txtRange.Length
        If txtRange.Characters(i, 1).Font.Color.Type = msoColorTypeRGB _
           Or txtRange.Characters(i, 1).Font.Color.Type = msoColorTypeScheme Then
            AddColor txtRange.Characters(i, 1).Font.Color.RGB, "Font Color"
        End If
    Next i
    
    On Error GoTo 0
End Sub

Private Sub ScanTableColors(tbl As table)
    Dim r As Long, c As Long
    
    On Error Resume Next
    
    For r = 1 To tbl.rows.count
        For c = 1 To tbl.Columns.count
                If tbl.cell(r, c).shape.Fill.visible Then
                    If tbl.cell(r, c).shape.Fill.ForeColor.Type = msoColorTypeRGB _
                       Or tbl.cell(r, c).shape.Fill.ForeColor.Type = msoColorTypeScheme Then
                        AddColor tbl.cell(r, c).shape.Fill.ForeColor.RGB, "Table Cell Fill"
                    End If
                End If
            
            ScanCellBorders tbl.cell(r, c)
            
            If tbl.cell(r, c).shape.HasTextFrame Then
                If tbl.cell(r, c).shape.TextFrame.HasText Then
                    ScanTextColors tbl.cell(r, c).shape.TextFrame.textRange
                End If
            End If
        Next c
    Next r
    
    On Error GoTo 0
End Sub

Private Sub ScanCellBorders(cell As cell)
    On Error Resume Next
    
        If cell.Borders(ppBorderTop).visible Then
            If cell.Borders(ppBorderTop).ForeColor.Type = msoColorTypeRGB _
               Or cell.Borders(ppBorderTop).ForeColor.Type = msoColorTypeScheme Then
                AddColor cell.Borders(ppBorderTop).ForeColor.RGB, "Table Border"
            End If
        End If
    
If cell.Borders(ppBorderBottom).visible Then
    If cell.Borders(ppBorderBottom).ForeColor.Type = msoColorTypeRGB _
       Or cell.Borders(ppBorderBottom).ForeColor.Type = msoColorTypeScheme Then
        AddColor cell.Borders(ppBorderBottom).ForeColor.RGB, "Table Border"
    End If
End If
    
If cell.Borders(ppBorderLeft).visible Then
    If cell.Borders(ppBorderLeft).ForeColor.Type = msoColorTypeRGB _
       Or cell.Borders(ppBorderLeft).ForeColor.Type = msoColorTypeScheme Then
        AddColor cell.Borders(ppBorderLeft).ForeColor.RGB, "Table Border"
    End If
End If
    
If cell.Borders(ppBorderRight).visible Then
    If cell.Borders(ppBorderRight).ForeColor.Type = msoColorTypeRGB _
       Or cell.Borders(ppBorderRight).ForeColor.Type = msoColorTypeScheme Then
        AddColor cell.Borders(ppBorderRight).ForeColor.RGB, "Table Border"
    End If
End If
    
    On Error GoTo 0
End Sub

Private Sub ScanChartColors(cht As Chart)
    On Error Resume Next
    
    Dim ser As Series
    Dim pt As Point
    Dim i As Long
    
    For Each ser In cht.SeriesCollection
        If ser.Format.Fill.visible Then
            If ser.Format.Fill.ForeColor.Type = msoColorTypeRGB _
               Or ser.Format.Fill.ForeColor.Type = msoColorTypeScheme Then
                AddColor ser.Format.Fill.ForeColor.RGB, "Chart Series"
            End If
        End If
        
         For i = 1 To ser.Points.count
            Set pt = ser.Points(i)
        If pt.Format.Fill.visible Then
            If pt.Format.Fill.ForeColor.Type = msoColorTypeRGB _
               Or pt.Format.Fill.ForeColor.Type = msoColorTypeScheme Then
                AddColor pt.Format.Fill.ForeColor.RGB, "Chart Point"
            End If
        End If
        Next i
    Next ser
    
    On Error GoTo 0
End Sub

Private Sub AddColor(rgbValue As Long, usageType As String)
    Dim r As Long, g As Long, b As Long
    
    r = rgbValue Mod 256
    g = (rgbValue \ 256) Mod 256
    b = (rgbValue \ 65536) Mod 256
    
    colorCount = colorCount + 1
    
    If colorCount > UBound(colors) Then
        ReDim Preserve colors(1 To colorCount + 100) As ColorInfo
    End If
    
    colors(colorCount).RGB = rgbValue
    colors(colorCount).RedValue = r
    colors(colorCount).GreenValue = g
    colors(colorCount).BlueValue = b
    colors(colorCount).HexValue = RGBToHex(r, g, b)
    colors(colorCount).UsageCount = 1
    colors(colorCount).UsageTypes = usageType
End Sub

Private Sub ConsolidateColors()
    Dim i As Long, j As Long
    Dim uniqueColors() As ColorInfo
    Dim uniqueCount As Long
    Dim found As Boolean
    
    If colorCount = 0 Then Exit Sub
    
    ReDim uniqueColors(1 To colorCount) As ColorInfo
    uniqueCount = 0
    
    For i = 1 To colorCount
        found = False
        
        For j = 1 To uniqueCount
            If colors(i).RGB = uniqueColors(j).RGB Then
                uniqueColors(j).UsageCount = uniqueColors(j).UsageCount + 1
                
                If InStr(uniqueColors(j).UsageTypes, colors(i).UsageTypes) = 0 Then
                    uniqueColors(j).UsageTypes = uniqueColors(j).UsageTypes & ", " & colors(i).UsageTypes
                End If
                
                found = True
                Exit For
            End If
        Next j
        
        If Not found Then
            uniqueCount = uniqueCount + 1
            uniqueColors(uniqueCount) = colors(i)
        End If
    Next i
    
    SortColorsByUsage uniqueColors, uniqueCount
    
    ReDim colors(1 To uniqueCount) As ColorInfo
    For i = 1 To uniqueCount
        colors(i) = uniqueColors(i)
    Next i
    colorCount = uniqueCount
End Sub

Private Sub SortColorsByUsage(arr() As ColorInfo, count As Long)
    Dim i As Long, j As Long
    Dim temp As ColorInfo
    
    For i = 1 To count - 1
        For j = i + 1 To count
            If arr(i).UsageCount < arr(j).UsageCount Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

Public Function RGBToHex(r As Long, g As Long, b As Long) As String
    RGBToHex = "#" & right("0" & Hex(r), 2) & right("0" & Hex(g), 2) & right("0" & Hex(b), 2)
End Function

Public Sub ReplaceColor(oldRGB As Long, newRGB As Long, scope As String)
    Dim sld As Slide
    Dim shp As shape
    Dim replaceCount As Long
    
    On Error Resume Next
    
   
    If scope = "selected slides" Then
        For Each sld In ActiveWindow.Selection.SlideRange
            replaceCount = replaceCount + ReplaceSlideBackgroundColor(sld, oldRGB, newRGB)
            
            For Each shp In sld.Shapes
                replaceCount = replaceCount + ReplaceShapeColor(shp, oldRGB, newRGB)
            Next shp
        Next sld
    Else
        For Each sld In ActivePresentation.Slides
            replaceCount = replaceCount + ReplaceSlideBackgroundColor(sld, oldRGB, newRGB)
            
            For Each shp In sld.Shapes
                replaceCount = replaceCount + ReplaceShapeColor(shp, oldRGB, newRGB)
            Next shp
        Next sld
    End If
    
  
    MsgBox "Replaced " & replaceCount & " instance(s) of the color in " & scope & ".", vbInformation, "Color Replacement Complete"
    
    On Error GoTo 0
End Sub

Private Function ReplaceSlideBackgroundColor(sld As Slide, oldRGB As Long, newRGB As Long) As Long
    Dim count As Long
    count = 0
    
    On Error Resume Next
    
        If sld.Background.Fill.visible Then
            If sld.Background.Fill.ForeColor.Type = msoColorTypeRGB _
               Or sld.Background.Fill.ForeColor.Type = msoColorTypeScheme Then
                If sld.Background.Fill.ForeColor.RGB = oldRGB Then
                    sld.Background.Fill.ForeColor.RGB = newRGB
                    count = count + 1
                End If
            End If
        End If
    
    ReplaceSlideBackgroundColor = count
    On Error GoTo 0
End Function

Private Function ReplaceShapeColor(shp As shape, oldRGB As Long, newRGB As Long) As Long
    Dim count As Long
    count = 0
    
    On Error Resume Next
    
    If shp.Type = msoGroup Then
        Dim subShape As shape
        For Each subShape In shp.GroupItems
            count = count + ReplaceShapeColor(subShape, oldRGB, newRGB)
        Next subShape
        ReplaceShapeColor = count
        Exit Function
    End If
    
        If shp.Fill.visible Then
            If shp.Fill.Type = msoFillSolid Then
                If shp.Fill.ForeColor.Type = msoColorTypeRGB _
                   Or shp.Fill.ForeColor.Type = msoColorTypeScheme Then
                    If shp.Fill.ForeColor.RGB = oldRGB Then
                        shp.Fill.ForeColor.RGB = newRGB
                        count = count + 1
                    End If
                End If
        ElseIf shp.Fill.Type = msoFillGradient Then
            Dim i As Long
            For i = 1 To shp.Fill.GradientStops.count
                If shp.Fill.GradientStops(i).Color.Type = msoColorTypeRGB _
                   Or shp.Fill.GradientStops(i).Color.Type = msoColorTypeScheme Then
                    If shp.Fill.GradientStops(i).Color.RGB = oldRGB Then
                        shp.Fill.GradientStops(i).Color.RGB = newRGB
                        count = count + 1
                    End If
                End If
            Next i
        End If
    End If

    If shp.Line.visible Then
        If shp.Line.ForeColor.Type = msoColorTypeRGB _
           Or shp.Line.ForeColor.Type = msoColorTypeScheme Then
            If shp.Line.ForeColor.RGB = oldRGB Then
                shp.Line.ForeColor.RGB = newRGB
                count = count + 1
            End If
        End If
    End If
    
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            count = count + ReplaceTextColor(shp.TextFrame.textRange, oldRGB, newRGB)
        End If
    End If
    
    If shp.HasTable Then
        count = count + ReplaceTableColors(shp.table, oldRGB, newRGB)
    End If
    
    If shp.HasChart Then
        count = count + ReplaceChartColors(shp.Chart, oldRGB, newRGB)
    End If
    
    ReplaceShapeColor = count
    On Error GoTo 0
End Function

Private Function ReplaceTextColor(txtRange As textRange, oldRGB As Long, newRGB As Long) As Long
    Dim i As Long
    Dim count As Long
    count = 0
    
    On Error Resume Next
    
    For i = 1 To txtRange.Length
        If txtRange.Characters(i, 1).Font.Color.Type = msoColorTypeRGB _
           Or txtRange.Characters(i, 1).Font.Color.Type = msoColorTypeScheme Then
            If txtRange.Characters(i, 1).Font.Color.RGB = oldRGB Then
                txtRange.Characters(i, 1).Font.Color.RGB = newRGB
                count = count + 1
            End If
        End If
    Next i
        
    ReplaceTextColor = count
    On Error GoTo 0
End Function

Private Function ReplaceTableColors(tbl As table, oldRGB As Long, newRGB As Long) As Long
    Dim r As Long, c As Long
    Dim count As Long
    count = 0
    
    On Error Resume Next
    
    For r = 1 To tbl.rows.count
        For c = 1 To tbl.Columns.count
            If tbl.cell(r, c).shape.Fill.visible Then
            If tbl.cell(r, c).shape.Fill.ForeColor.Type = msoColorTypeRGB _
               Or tbl.cell(r, c).shape.Fill.ForeColor.Type = msoColorTypeScheme Then
                If tbl.cell(r, c).shape.Fill.ForeColor.RGB = oldRGB Then
                    tbl.cell(r, c).shape.Fill.ForeColor.RGB = newRGB
                    count = count + 1
                End If
            End If
        End If

            count = count + ReplaceCellBorderColors(tbl.cell(r, c), oldRGB, newRGB)

            If tbl.cell(r, c).shape.HasTextFrame Then
                If tbl.cell(r, c).shape.TextFrame.HasText Then
                    count = count + ReplaceTextColor(tbl.cell(r, c).shape.TextFrame.textRange, oldRGB, newRGB)
                End If
            End If
        Next c
    Next r
    
    ReplaceTableColors = count
    On Error GoTo 0
End Function

Private Function ReplaceCellBorderColors(cell As cell, oldRGB As Long, newRGB As Long) As Long
    Dim count As Long
    count = 0
    
    On Error Resume Next
    
If cell.Borders(ppBorderTop).visible Then
    If cell.Borders(ppBorderTop).ForeColor.Type = msoColorTypeRGB _
       Or cell.Borders(ppBorderTop).ForeColor.Type = msoColorTypeScheme Then
        If cell.Borders(ppBorderTop).ForeColor.RGB = oldRGB Then
            cell.Borders(ppBorderTop).ForeColor.RGB = newRGB
            count = count + 1
        End If
    End If
End If
    
If cell.Borders(ppBorderBottom).visible Then
    If cell.Borders(ppBorderBottom).ForeColor.Type = msoColorTypeRGB _
       Or cell.Borders(ppBorderBottom).ForeColor.Type = msoColorTypeScheme Then
        If cell.Borders(ppBorderBottom).ForeColor.RGB = oldRGB Then
            cell.Borders(ppBorderBottom).ForeColor.RGB = newRGB
            count = count + 1
        End If
    End If
End If
    
If cell.Borders(ppBorderLeft).visible Then
    If cell.Borders(ppBorderLeft).ForeColor.Type = msoColorTypeRGB _
       Or cell.Borders(ppBorderLeft).ForeColor.Type = msoColorTypeScheme Then
        If cell.Borders(ppBorderLeft).ForeColor.RGB = oldRGB Then
            cell.Borders(ppBorderLeft).ForeColor.RGB = newRGB
            count = count + 1
        End If
    End If
End If
    
If cell.Borders(ppBorderRight).visible Then
    If cell.Borders(ppBorderRight).ForeColor.Type = msoColorTypeRGB _
       Or cell.Borders(ppBorderRight).ForeColor.Type = msoColorTypeScheme Then
        If cell.Borders(ppBorderRight).ForeColor.RGB = oldRGB Then
            cell.Borders(ppBorderRight).ForeColor.RGB = newRGB
            count = count + 1
        End If
    End If
End If
    
    ReplaceCellBorderColors = count
    On Error GoTo 0
End Function

Private Function ReplaceChartColors(cht As Chart, oldRGB As Long, newRGB As Long) As Long
    Dim count As Long
    Dim ser As Series
    Dim pt As Point
    Dim i As Long
    
    count = 0
    
    On Error Resume Next
    
    For Each ser In cht.SeriesCollection
        If ser.Format.Fill.visible Then
            If ser.Format.Fill.ForeColor.Type = msoColorTypeRGB _
               Or ser.Format.Fill.ForeColor.Type = msoColorTypeScheme Then
                If ser.Format.Fill.ForeColor.RGB = oldRGB Then
                    ser.Format.Fill.ForeColor.RGB = newRGB
                    count = count + 1
                End If
            End If
        End If
        
        For i = 1 To ser.Points.count
            Set pt = ser.Points(i)
            If pt.Format.Fill.visible Then
                If pt.Format.Fill.ForeColor.Type = msoColorTypeRGB _
                   Or pt.Format.Fill.ForeColor.Type = msoColorTypeScheme Then
                    If pt.Format.Fill.ForeColor.RGB = oldRGB Then
                        pt.Format.Fill.ForeColor.RGB = newRGB
                        count = count + 1
                    End If
                End If
        End If
        Next i
    Next ser
    
    ReplaceChartColors = count
    On Error GoTo 0
End Function

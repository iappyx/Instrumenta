Attribute VB_Name = "ModuleStyleSheets"
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

Public Const STYLE_H1_SIZE As Long = 40
Public Const STYLE_H2_SIZE As Long = 32
Public Const STYLE_H3_SIZE As Long = 26
Public Const STYLE_PARAGRAPH_SIZE As Long = 20
Public Const STYLE_QUOTE_SIZE As Long = 20

Public Const STYLE_H1_BOLD As Boolean = True
Public Const STYLE_H1_ITALIC As Boolean = False

Public Const STYLE_H2_BOLD As Boolean = True
Public Const STYLE_H2_ITALIC As Boolean = False

Public Const STYLE_H3_BOLD As Boolean = True
Public Const STYLE_H3_ITALIC As Boolean = False

Public Const STYLE_PARAGRAPH_BOLD As Boolean = False
Public Const STYLE_PARAGRAPH_ITALIC As Boolean = False

Public Const STYLE_QUOTE_BOLD As Boolean = False
Public Const STYLE_QUOTE_ITALIC As Boolean = True

Public Const STYLE_CUSTOM1_SIZE As Long = 24
Public Const STYLE_CUSTOM1_BOLD As Boolean = True
Public Const STYLE_CUSTOM1_ITALIC As Boolean = False

Public Const STYLE_CUSTOM2_SIZE As Long = 18
Public Const STYLE_CUSTOM2_BOLD As Boolean = False
Public Const STYLE_CUSTOM2_ITALIC As Boolean = True

Public Const STYLE_CUSTOM3_SIZE As Long = 16
Public Const STYLE_CUSTOM3_BOLD As Boolean = False
Public Const STYLE_CUSTOM3_ITALIC As Boolean = False

Public Const STYLE_CUSTOM4_SIZE As Long = 22
Public Const STYLE_CUSTOM4_BOLD As Boolean = True
Public Const STYLE_CUSTOM4_ITALIC As Boolean = True

Public Const STYLE_CUSTOM5_SIZE As Long = 28
Public Const STYLE_CUSTOM5_BOLD As Boolean = True
Public Const STYLE_CUSTOM5_ITALIC As Boolean = False

Public Const STYLE_COLUMN1_LEFT As Long = 50
Public Const STYLE_COLUMN2_LEFT As Long = 500
Public Const STYLE_COLUMN_TOP As Long = 50
Public Const STYLE_SHAPE_WIDTH As Long = 400
Public Const STYLE_SHAPE_HEIGHT As Long = 60
Public Const STYLE_SHAPE_SPACING As Long = 20


Function GetCurrentSlideMaster() As Object
    
    Set GetCurrentSlideMaster = Nothing
    
    On Error Resume Next
    
    Set GetCurrentSlideMaster = ActiveWindow.Selection.SlideRange(1).Design.SlideMaster
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Please exit Master View and select a slide first.", vbExclamation
        Exit Function
    End If
    
    On Error GoTo 0
    
    If GetCurrentSlideMaster Is Nothing Then
        MsgBox "Please select a slide or shape first.", vbExclamation
    End If
End Function


Sub CreateStyleSheetLayout()

    Dim sm As Object
    Dim layout As CustomLayout
    Dim answer As VbMsgBoxResult

    Set sm = GetCurrentSlideMaster()
    
    If sm Is Nothing Then
        MsgBox "Please select a slide or enter Slide Master view first.", vbExclamation
        Exit Sub
    End If

    For Each layout In sm.CustomLayouts
        If layout.name = "InstrumentaStylesheet" Then
            
            answer = MsgBox("Instrumenta stylesheet already exists on this master." & vbCrLf & _
                            "Do you want to delete it and recreate it?", _
                            vbYesNo + vbQuestion, "Recreate Stylesheet?")
            
            If answer = vbNo Then Exit Sub
            
            layout.delete
            Exit For
        End If
    Next layout

    Call CreateStyleSheetOnMaster(sm)

    MsgBox "Instrumenta stylesheet layout created on the current slide master."

End Sub

Sub CreateStyleShape(layout As CustomLayout, name As String, text As String, _
                     ByRef topPos As Single, fontSize As Long, _
                     isBold As Boolean, isItalic As Boolean, leftPos As Single)

    Dim shp As shape
    Set shp = layout.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, topPos, STYLE_SHAPE_WIDTH, STYLE_SHAPE_HEIGHT)

    shp.name = name

    shp.TextFrame2.AutoSize = msoAutoSizeNone

    shp.TextFrame2.textRange.text = text

    With shp.TextFrame2.textRange.Font
        .name = "Segoe UI"
        .Size = fontSize
        .Bold = IIf(isBold, msoTrue, msoFalse)
        .Italic = IIf(isItalic, msoTrue, msoFalse)
    End With

    topPos = topPos + STYLE_SHAPE_HEIGHT + STYLE_SHAPE_SPACING

End Sub


Function FindStylesheet(sm As Object) As CustomLayout
    
    Dim layout As CustomLayout
    
    For Each layout In sm.CustomLayouts
        If layout.name = "InstrumentaStylesheet" Then
            Set FindStylesheet = layout
            Exit Function
        End If
    Next layout
    
    Set FindStylesheet = Nothing
End Function


Sub ApplyTextStyle(styleName As String)

    Dim sm As Object
    Dim layout As CustomLayout
    Dim styleShp As shape
    Dim sel As Selection
    Dim stylesheet As CustomLayout

    Set sel = ActiveWindow.Selection
    
    Set sm = GetCurrentSlideMaster()
    
    If sm Is Nothing Then
        MsgBox "Please select a slide or shape first.", vbExclamation
        Exit Sub
    End If

    Set stylesheet = FindStylesheet(sm)
    
    If stylesheet Is Nothing Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox( _
            "No Instrumenta stylesheet found on this slide master." & vbCrLf & vbCrLf & _
            "Do you want to create one now?", _
            vbYesNo + vbQuestion, _
            "Create Stylesheet?")
        
        If answer = vbYes Then
            CreateStyleSheetLayout
        End If
    
        Exit Sub
    End If

    For Each styleShp In stylesheet.Shapes
        If styleShp.name = styleName Then

            If sel.Type = ppSelectionShapes Then
                styleShp.PickUp
                sel.ShapeRange(1).Apply
            
                sel.ShapeRange(1).Tags.Add "InstrumentaStyle", styleName
            
                Exit Sub
            End If

            If sel.Type = ppSelectionText Then
                ApplyTextFormatting sel.TextRange2, styleShp.TextFrame2.textRange
                Exit Sub
            End If

        End If
    Next styleShp

    MsgBox "Text style not found: " & styleName

End Sub


Sub ApplyTextFormatting(target As TextRange2, source As TextRange2)

    With target.Font
        .name = source.Font.name
        .Size = source.Font.Size
        .Bold = source.Font.Bold
        .Italic = source.Font.Italic
        .UnderlineStyle = source.Font.UnderlineStyle
        .Fill.ForeColor.RGB = source.Font.Fill.ForeColor.RGB
        .BaselineOffset = source.Font.BaselineOffset
        .Kerning = source.Font.Kerning
        .spacing = source.Font.spacing
        .Caps = source.Font.Caps
        .Strike = source.Font.Strike

        .Glow.Radius = source.Font.Glow.Radius
        .Glow.color.RGB = source.Font.Glow.color.RGB
        .Reflection.Type = source.Font.Reflection.Type
    End With

    With target.ParagraphFormat
        .Alignment = source.ParagraphFormat.Alignment
        .FirstLineIndent = source.ParagraphFormat.FirstLineIndent
        .LeftIndent = source.ParagraphFormat.LeftIndent
        .RightIndent = source.ParagraphFormat.RightIndent
        .SpaceBefore = source.ParagraphFormat.SpaceBefore
        .SpaceAfter = source.ParagraphFormat.SpaceAfter

        On Error Resume Next
        .TextDirection = source.ParagraphFormat.TextDirection
        On Error GoTo 0
    End With

End Sub


Sub ApplyH1(): ApplyTextStyle "Style_H1": End Sub
Sub ApplyH2(): ApplyTextStyle "Style_H2": End Sub
Sub ApplyH3(): ApplyTextStyle "Style_H3": End Sub
Sub ApplyParagraph(): ApplyTextStyle "Style_Paragraph": End Sub
Sub ApplyQuote(): ApplyTextStyle "Style_Quote": End Sub
Sub ApplyCustom1(): ApplyTextStyle "Style_Custom1": End Sub
Sub ApplyCustom2(): ApplyTextStyle "Style_Custom2": End Sub
Sub ApplyCustom3(): ApplyTextStyle "Style_Custom3": End Sub
Sub ApplyCustom4(): ApplyTextStyle "Style_Custom4": End Sub
Sub ApplyCustom5(): ApplyTextStyle "Style_Custom5": End Sub


Sub OpenStyleSheet()

    Dim sm As Object
    Dim stylesheet As CustomLayout

    Set sm = GetCurrentSlideMaster()
    
    If sm Is Nothing Then
        MsgBox "Please select a slide first.", vbExclamation
        Exit Sub
    End If

    Set stylesheet = FindStylesheet(sm)
    
    If stylesheet Is Nothing Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox( _
            "No Instrumenta stylesheet found on this slide master." & vbCrLf & vbCrLf & _
            "Do you want to create one now?", _
            vbYesNo + vbQuestion, _
            "Create Stylesheet?")
        
        If answer = vbYes Then
            CreateStyleSheetLayout
            DoEvents
            Set stylesheet = FindStylesheet(sm)
        
            If stylesheet Is Nothing Then
                Exit Sub
            End If
        
        Else
        Exit Sub
        End If

    
End If


    ActiveWindow.ViewType = ppViewSlideMaster

    sm.CustomLayouts(1).Select

    stylesheet.Select

End Sub

Sub UpdateFullShapeStyles()

    Dim sm As Object
    Dim layout As CustomLayout
    Dim styleShp As shape
    Dim sld As Slide
    Dim shp As shape
    Dim styleName As String
    Dim stylesheet As CustomLayout
    Dim updatedCount As Long

    updatedCount = 0

    For Each sld In ActivePresentation.Slides
        
        Set sm = sld.Design.SlideMaster
        
        Set stylesheet = FindStylesheet(sm)
        
        If Not stylesheet Is Nothing Then
        
            For Each shp In sld.Shapes

                styleName = shp.Tags("InstrumentaStyle")
                If styleName <> "" Then

                    For Each styleShp In stylesheet.Shapes
                        If styleShp.name = styleName Then

                            styleShp.PickUp
                            shp.Apply
                            updatedCount = updatedCount + 1

                            Exit For
                        End If
                    Next styleShp

                End If

            Next shp
        
        End If
        
    Next sld

    MsgBox updatedCount & " shape(s) updated with their Instrumenta styles.", vbInformation, "Update Complete"

End Sub


Sub RemoveAllInstrumentaStyleTags()

    Dim sld As Slide
    Dim shp As shape
    Dim countRemoved As Long

    countRemoved = 0

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes

            If shp.Tags("InstrumentaStyle") <> "" Then
                shp.Tags.delete "InstrumentaStyle"
                countRemoved = countRemoved + 1
            End If

        Next shp
    Next sld

    MsgBox countRemoved & " Instrumenta style tag(s) removed.", vbInformation, "Cleanup Complete"

End Sub


Sub RemoveInstrumentaStylesheet()

    Dim sm As Object
    Dim stylesheet As CustomLayout

    Set sm = GetCurrentSlideMaster()
    
    If sm Is Nothing Then
        MsgBox "Please select a slide or enter Slide Master view first.", vbExclamation
        Exit Sub
    End If

    Set stylesheet = FindStylesheet(sm)
    
    If stylesheet Is Nothing Then
        MsgBox "No Instrumenta stylesheet layout found on this slide master.", vbInformation, "Nothing to Remove"
        Exit Sub
    End If

    stylesheet.delete

    MsgBox "Instrumenta stylesheet layout has been removed from this slide master.", _
           vbInformation, "Stylesheet Removed"

End Sub

Sub CreateStyleSheetOnAllMasters()
    
    Dim d As Design
    Dim sm As Object
    Dim createdCount As Long
    Dim skippedCount As Long
    Dim layout As CustomLayout
    Dim hasStylesheet As Boolean
    
    createdCount = 0
    skippedCount = 0
    
    For Each d In ActivePresentation.Designs
        Set sm = d.SlideMaster
        
        hasStylesheet = False
        For Each layout In sm.CustomLayouts
            If layout.name = "InstrumentaStylesheet" Then
                hasStylesheet = True
                Exit For
            End If
        Next layout
        
        If hasStylesheet Then
            skippedCount = skippedCount + 1
        Else
            Call CreateStyleSheetOnMaster(sm)
            createdCount = createdCount + 1
        End If
        
    Next d
    
    MsgBox "Created stylesheets on " & createdCount & " master(s)." & vbCrLf & _
           "Skipped " & skippedCount & " master(s) that already had stylesheets.", _
           vbInformation, "Batch Creation Complete"
           
End Sub


Sub CreateStyleSheetOnMaster(sm As Object)
    
    Dim layout As CustomLayout
    Dim topPos As Single
    Dim warn As shape
    Dim slideWidth As Single
    Dim slideHeight As Single
    Dim marginBottom As Single
    
    Set layout = sm.CustomLayouts(sm.CustomLayouts.count).Duplicate
    layout.name = "InstrumentaStylesheet"

    Do While layout.Shapes.count > 0
        layout.Shapes(1).delete
    Loop

    topPos = STYLE_COLUMN_TOP
    CreateStyleShape layout, "Style_H1", "Heading 1", topPos, STYLE_H1_SIZE, STYLE_H1_BOLD, STYLE_H1_ITALIC, STYLE_COLUMN1_LEFT
    CreateStyleShape layout, "Style_H2", "Heading 2", topPos, STYLE_H2_SIZE, STYLE_H2_BOLD, STYLE_H2_ITALIC, STYLE_COLUMN1_LEFT
    CreateStyleShape layout, "Style_H3", "Heading 3", topPos, STYLE_H3_SIZE, STYLE_H3_BOLD, STYLE_H3_ITALIC, STYLE_COLUMN1_LEFT
    CreateStyleShape layout, "Style_Paragraph", "Paragraph", topPos, STYLE_PARAGRAPH_SIZE, STYLE_PARAGRAPH_BOLD, STYLE_PARAGRAPH_ITALIC, STYLE_COLUMN1_LEFT
    CreateStyleShape layout, "Style_Quote", "Quote", topPos, STYLE_QUOTE_SIZE, STYLE_QUOTE_BOLD, STYLE_QUOTE_ITALIC, STYLE_COLUMN1_LEFT
    
    topPos = STYLE_COLUMN_TOP
    CreateStyleShape layout, "Style_Custom1", "Custom 1", topPos, STYLE_CUSTOM1_SIZE, STYLE_CUSTOM1_BOLD, STYLE_CUSTOM1_ITALIC, STYLE_COLUMN2_LEFT
    CreateStyleShape layout, "Style_Custom2", "Custom 2", topPos, STYLE_CUSTOM2_SIZE, STYLE_CUSTOM2_BOLD, STYLE_CUSTOM2_ITALIC, STYLE_COLUMN2_LEFT
    CreateStyleShape layout, "Style_Custom3", "Custom 3", topPos, STYLE_CUSTOM3_SIZE, STYLE_CUSTOM3_BOLD, STYLE_CUSTOM3_ITALIC, STYLE_COLUMN2_LEFT
    CreateStyleShape layout, "Style_Custom4", "Custom 4", topPos, STYLE_CUSTOM4_SIZE, STYLE_CUSTOM4_BOLD, STYLE_CUSTOM4_ITALIC, STYLE_COLUMN2_LEFT
    CreateStyleShape layout, "Style_Custom5", "Custom 5", topPos, STYLE_CUSTOM5_SIZE, STYLE_CUSTOM5_BOLD, STYLE_CUSTOM5_ITALIC, STYLE_COLUMN2_LEFT

    slideWidth = sm.width
    slideHeight = sm.height
    marginBottom = 10

    Set warn = layout.Shapes.AddTextbox( _
        orientation:=msoTextOrientationHorizontal, _
        left:=0, _
        Top:=slideHeight - marginBottom - 80, _
        width:=slideWidth, _
        height:=80)

    With warn
        .name = "InstrumentaWarning"
        .TextFrame2.textRange.text = _
            "DO NOT USE THIS LAYOUT" & vbCrLf & _
            "This layout is for Instrumenta stylesheet definitions only."

        With .TextFrame2.textRange.Font
            .name = "Segoe UI"
            .Size = 20
            .Bold = msoTrue
            .Italic = msoFalse
            .Fill.ForeColor.RGB = RGB(200, 0, 0)
        End With

        .TextFrame2.textRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.marginBottom = 0
    End With
    
End Sub

Sub RemoveStyleSheetsFromAllMasters()
    
    Dim d As Design
    Dim sm As Object
    Dim removedCount As Long
    Dim skippedCount As Long
    Dim stylesheet As CustomLayout
    
    removedCount = 0
    skippedCount = 0
    
    For Each d In ActivePresentation.Designs
        Set sm = d.SlideMaster
        
        Set stylesheet = FindStylesheet(sm)
        
        If stylesheet Is Nothing Then
            skippedCount = skippedCount + 1
        Else
            stylesheet.delete
            removedCount = removedCount + 1
        End If
        
    Next d
    
    MsgBox "Removed stylesheets from " & removedCount & " master(s)." & vbCrLf & _
           "Skipped " & skippedCount & " master(s) with no stylesheets.", _
           vbInformation, "Batch Removal Complete"
           
End Sub


Sub ExportStylesToPPTX()

    Dim sm As Object
    Dim stylesheet As CustomLayout
    Dim tempPres As Presentation
    Dim tempSlide As Slide
    Dim shp As shape
    Dim exportPath As String
    Dim DotPosition As Long

    Set sm = GetCurrentSlideMaster()
    If sm Is Nothing Then
        Exit Sub
    End If

    Set stylesheet = FindStylesheet(sm)
    If stylesheet Is Nothing Then
        MsgBox "InstrumentaStylesheet not found on this master. Nothing to export.", vbExclamation
        Exit Sub
    End If

    #If Mac Then
        exportPath = MacSaveAsDialog("InstrumentaStylesheet.pptx")
    #Else
        Dim exportFileDialog As FileDialog
        Set exportFileDialog = Application.FileDialog(msoFileDialogSaveAs)
        exportFileDialog.InitialFileName = "InstrumentaStylesheet.pptx"
        
        If exportFileDialog.Show = -1 Then
            exportPath = exportFileDialog.SelectedItems(1)
        Else
            Exit Sub
        End If
    #End If
    
    DotPosition = InStrRev(exportPath, ".")
    If DotPosition > 0 Then
        exportPath = left(exportPath, DotPosition - 1) & ".pptx"
    Else
        exportPath = exportPath & ".pptx"
    End If

    Set tempPres = Presentations.Add(msoFalse)
    Set tempSlide = tempPres.Slides.Add(1, ppLayoutBlank)

    For Each shp In stylesheet.Shapes
        shp.Copy
        tempSlide.Shapes.Paste
    Next shp

    tempPres.SaveAs exportPath
    tempPres.Close

    MsgBox "Stylesheet exported to:" & vbCrLf & exportPath, vbInformation

End Sub

Sub ImportStylesFromPPTX()

    Dim sm As Object
    Dim stylesheet As CustomLayout
    Dim importPres As Presentation
    Dim importSlide As Slide
    Dim shp As shape
    Dim importPath As String

    Set sm = GetCurrentSlideMaster()
    If sm Is Nothing Then
        Exit Sub
    End If

    Set stylesheet = FindStylesheet(sm)
    
If stylesheet Is Nothing Then
    Dim answer As VbMsgBoxResult
    answer = MsgBox( _
        "No Instrumenta stylesheet found on this slide master." & vbCrLf & vbCrLf & _
        "You need to create one first before you can import. " & vbCrLf & vbCrLf & _
        "Do you want to create one now?", _
        vbYesNo + vbQuestion, _
        "Create Stylesheet?")
    
    If answer = vbYes Then
        CreateStyleSheetLayout
        DoEvents
        Set stylesheet = FindStylesheet(sm)
        
        If stylesheet Is Nothing Then
        Exit Sub
        End If
        
        
    Else

    Exit Sub
    End If
End If


    #If Mac Then
       
           importPath = MacFileDialog("/")
            
            If importPath = "" Then
                MsgBox "No file selected."
                Exit Sub
            End If
        
    #Else
        Dim importFileDialog As FileDialog
        Set importFileDialog = Application.FileDialog(msoFileDialogFilePicker)
        importFileDialog.Title = "Import Instrumenta Stylesheet"
        importFileDialog.Filters.Clear
        importFileDialog.Filters.Add "PowerPoint Files", "*.pptx"
        
        If importFileDialog.Show = -1 Then
            importPath = importFileDialog.SelectedItems(1)
        Else
            Exit Sub
        End If
    #End If
    
    If importPath = "" Then Exit Sub
    
    Set importPres = Presentations.Open(importPath, WithWindow:=msoFalse)
    Set importSlide = importPres.Slides(1)
    
    Do While stylesheet.Shapes.count > 0
        stylesheet.Shapes(1).delete
    Loop

    For Each shp In importSlide.Shapes
        shp.Copy
        stylesheet.Shapes.Paste
    Next shp

    importPres.Close

    MsgBox "Stylesheet imported successfully.", vbInformation

End Sub

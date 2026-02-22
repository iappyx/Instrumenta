Attribute VB_Name = "ModuleSlideGrader"
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

Public Type SlideIssue
    slideIndex  As Long
    SlideNumber As Long
    Severity    As String
    checkName   As String
    Description As String
    fixKey      As String
    auxData     As String
    IsFixed     As Boolean
End Type

Public Type GraderCheckConfig
    
    T1_MissingTitle     As Boolean
    T2_ActionTitle      As Boolean
    T3_TitleFontSize    As Boolean
    T4_TitleFontFace    As Boolean
    T5_TitleEndsPeriod  As Boolean
    T6_DuplicateTitles  As Boolean
    
    Q1_DoubleSpaces     As Boolean
    Q2_TrailingSpaces   As Boolean
    Q3_PlaceholderText  As Boolean
    Q4_EmptyTextBox     As Boolean
    Q5_ExcessiveWords   As Boolean
    Q6_BulletPunctuation As Boolean
    Q7_Strikethrough    As Boolean
    Q8_DoublePunct      As Boolean
    
    F1_FontFamilyMix    As Boolean
    F2_FontSizeMix      As Boolean
    F3_FontTooSmall     As Boolean
    F4_UnderlineText    As Boolean
    F5_AllCapsBody      As Boolean
    F6_BodyColorIncons  As Boolean
    F7_InvisibleText    As Boolean
    
    L1_TextOverflow     As Boolean
    L2_OutsideBounds    As Boolean
    L3_OverlappingShapes As Boolean
    L4_InvisibleShape   As Boolean
    L5_RotatedTextBox   As Boolean
    L6_TooManyShapes    As Boolean
    L7_NoContent        As Boolean
    L8_IsolatedFloat    As Boolean
    
    C1_TooManyColors    As Boolean
    C2_FillColorIncons  As Boolean
    C3_AccentIncons     As Boolean
    C4_LowContrast      As Boolean
    
    D1_TableHeader      As Boolean
    D2_EmptyTableCells  As Boolean
    D3_HiddenSlides     As Boolean
    D4_NotesDraft       As Boolean
    
    MaxFontFamilies     As Integer
    MaxFontSizes        As Integer
    MaxShapesPerSlide   As Integer
    MaxBodyWords        As Integer
    MinFontSizePt       As Single
    MaxColorsPerSlide   As Integer
End Type

Private m_Issues()         As SlideIssue
Private m_IssueCount       As Long
Private m_HighlightSlideIdx As Long


Private m_TitleFontSizesArr()  As Single
Private m_TitleFontNamesArr()  As String
Private m_TitleLayoutNames()   As String
Private m_TitleCount           As Long
Private m_BodyColorsArr()      As Long
Private m_BodyColorCount       As Long
Private m_AllTitles()          As String
Private m_AllTitleSlideIdx()   As Long

Public Sub ShowSlideGrader()
    If Application.Presentations.count = 0 Then
        MsgBox "Please open a presentation first.", vbExclamation, "Slide Grader"
        Exit Sub
    End If
    SlideGraderForm.Show vbModeless
End Sub

Public Function GetIssueCount() As Long
    GetIssueCount = m_IssueCount
End Function

Public Function GetIssue(idx As Long) As SlideIssue
    If idx >= 1 And idx <= m_IssueCount Then
        GetIssue = m_Issues(idx)
    End If
End Function

Public Function GetDefaultConfig() As GraderCheckConfig
    Dim cfg As GraderCheckConfig
    
    cfg.T1_MissingTitle = True
    cfg.T2_ActionTitle = True
    cfg.T3_TitleFontSize = True
    cfg.T4_TitleFontFace = True
    cfg.T5_TitleEndsPeriod = True
    cfg.T6_DuplicateTitles = True
    cfg.Q1_DoubleSpaces = True
    cfg.Q2_TrailingSpaces = True
    cfg.Q3_PlaceholderText = True
    cfg.Q4_EmptyTextBox = True
    cfg.Q5_ExcessiveWords = True
    cfg.Q6_BulletPunctuation = True
    cfg.Q7_Strikethrough = True
    cfg.Q8_DoublePunct = True
    cfg.F1_FontFamilyMix = True
    cfg.F2_FontSizeMix = True
    cfg.F3_FontTooSmall = True
    cfg.F4_UnderlineText = True
    cfg.F5_AllCapsBody = True
    cfg.F6_BodyColorIncons = True
    cfg.F7_InvisibleText = True
    cfg.L1_TextOverflow = True
    cfg.L2_OutsideBounds = True
    cfg.L3_OverlappingShapes = True
    cfg.L4_InvisibleShape = True
    cfg.L5_RotatedTextBox = True
    cfg.L6_TooManyShapes = True
    cfg.L7_NoContent = True
    cfg.L8_IsolatedFloat = True
    cfg.C1_TooManyColors = True
    cfg.C2_FillColorIncons = True
    cfg.C3_AccentIncons = True
    cfg.C4_LowContrast = True
    cfg.D1_TableHeader = True
    cfg.D2_EmptyTableCells = True
    cfg.D3_HiddenSlides = True
    cfg.D4_NotesDraft = True
    
    cfg.MaxFontFamilies = 2
    cfg.MaxFontSizes = 4
    cfg.MaxShapesPerSlide = 20
    cfg.MaxBodyWords = 150
    cfg.MinFontSizePt = 10
    cfg.MaxColorsPerSlide = 5
    GetDefaultConfig = cfg
End Function

Public Sub RunSlideGrader(cfg As GraderCheckConfig)

    If Application.Presentations.count = 0 Then Exit Sub

    
    m_IssueCount = 0
    ReDim m_Issues(1 To 64)

    
    ComputeBaselines cfg

    
    Dim sld As PowerPoint.Slide
    For Each sld In ActivePresentation.Slides
        If cfg.T1_MissingTitle Or cfg.T2_ActionTitle Or cfg.T3_TitleFontSize _
           Or cfg.T4_TitleFontFace Or cfg.T5_TitleEndsPeriod Or cfg.T6_DuplicateTitles Then
            CheckTitles sld, cfg
        End If
        If cfg.Q1_DoubleSpaces Or cfg.Q2_TrailingSpaces Or cfg.Q3_PlaceholderText _
           Or cfg.Q4_EmptyTextBox Or cfg.Q5_ExcessiveWords Or cfg.Q6_BulletPunctuation _
           Or cfg.Q7_Strikethrough Or cfg.Q8_DoublePunct Then
            CheckTextQuality sld, cfg
        End If
        If cfg.F1_FontFamilyMix Or cfg.F2_FontSizeMix Or cfg.F3_FontTooSmall _
           Or cfg.F4_UnderlineText Or cfg.F5_AllCapsBody Or cfg.F6_BodyColorIncons _
           Or cfg.F7_InvisibleText Then
            CheckTypography sld, cfg
        End If
        If cfg.L1_TextOverflow Or cfg.L2_OutsideBounds Or cfg.L3_OverlappingShapes _
           Or cfg.L4_InvisibleShape Or cfg.L5_RotatedTextBox Or cfg.L6_TooManyShapes _
           Or cfg.L7_NoContent Or cfg.L8_IsolatedFloat Then
            CheckLayout sld, cfg
        End If
        If cfg.C1_TooManyColors Or cfg.C2_FillColorIncons Or cfg.C3_AccentIncons _
           Or cfg.C4_LowContrast Then
            CheckColors sld, cfg
        End If
        If cfg.D1_TableHeader Or cfg.D2_EmptyTableCells _
           Or cfg.D3_HiddenSlides Or cfg.D4_NotesDraft Then
            CheckTablesDeck sld, cfg
        End If
    Next sld

End Sub

Private Sub ComputeBaselines(cfg As GraderCheckConfig)

    Dim sld As PowerPoint.Slide
    Dim ttl As String
    Dim n As Long

    n = ActivePresentation.Slides.count

    ReDim m_TitleFontSizesArr(1 To n)
    ReDim m_TitleFontNamesArr(1 To n)
    ReDim m_TitleLayoutNames(1 To n)
    ReDim m_AllTitles(1 To n)
    ReDim m_AllTitleSlideIdx(1 To n)
    ReDim m_BodyColorsArr(1 To n * 10)

    m_TitleCount = 0
    m_BodyColorCount = 0

    Dim shp As PowerPoint.shape

    For Each sld In ActivePresentation.Slides

        
        Dim titleSize As Single
        Dim titleFace As String
        titleSize = 0
        titleFace = ""

        For Each shp In sld.shapes
            On Error Resume Next
            If shp.Type = msoPlaceholder Then
                If shp.PlaceholderFormat.Type = ppPlaceholderTitle Or _
                   shp.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
                    ttl = Trim(shp.TextFrame.textRange.text)
                    If Len(ttl) > 0 Then
                        titleSize = shp.TextFrame.textRange.Font.Size
                        titleFace = shp.TextFrame.textRange.Font.name
                        m_TitleCount = m_TitleCount + 1
                        m_TitleFontSizesArr(m_TitleCount) = titleSize
                        m_TitleFontNamesArr(m_TitleCount) = titleFace
                        m_TitleLayoutNames(m_TitleCount) = sld.CustomLayout.name
                        m_AllTitles(m_TitleCount) = LCase(ttl)
                        m_AllTitleSlideIdx(m_TitleCount) = sld.slideIndex
                    End If
                End If
            End If
            On Error GoTo 0
        Next shp

        
        For Each shp In sld.shapes
            On Error Resume Next
            If shp.Type <> msoPlaceholder And shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    Dim clr As Long
                    clr = shp.TextFrame.textRange.Font.color.RGB
                    If clr <> 0 And clr <> RGB(255, 255, 255) Then
                        m_BodyColorCount = m_BodyColorCount + 1
                        If m_BodyColorCount <= UBound(m_BodyColorsArr) Then
                            m_BodyColorsArr(m_BodyColorCount) = clr
                        End If
                    End If
                End If
            End If
            On Error GoTo 0
        Next shp

    Next sld

End Sub

Private Sub CheckTitles(sld As PowerPoint.Slide, cfg As GraderCheckConfig)

    Dim ttl As String
    ttl = GetSlideTitle(sld)

    
    If cfg.T1_MissingTitle Then
        If ttl = "" Then
            AddIssue sld.slideIndex, "Error", "Missing Title", "No title or empty title placeholder", "", GetTitleShapeName(sld)
            Exit Sub
        End If
    End If

    If ttl = "" Then Exit Sub

    
    If cfg.T2_ActionTitle Then
        Dim wc As Integer
        wc = CountWords(ttl)
        Dim flagged As Boolean
        Dim reason As String
        flagged = False

        If wc < 5 Then
            flagged = True
            reason = "Short title (" & wc & " word" & IIf(wc = 1, "", "s") & "): " & Chr(34) & left(ttl, 45) & IIf(Len(ttl) > 45, "...", "") & Chr(34)
        ElseIf wc < 7 And Not TitleHasVerb(ttl) Then
            flagged = True
            reason = "No insight verb (" & wc & " words): " & Chr(34) & left(ttl, 45) & IIf(Len(ttl) > 45, "...", "") & Chr(34)
        End If

        If flagged Then AddIssue sld.slideIndex, "Warning", "Action Title", reason, "", GetTitleShapeName(sld)
    End If

    
    If cfg.T3_TitleFontSize And m_TitleCount > 1 Then
        Dim modeSize As Single
        modeSize = GetModeSingleForLayout(m_TitleFontSizesArr, m_TitleLayoutNames, m_TitleCount, sld.CustomLayout.name)
        Dim thisSize As Single
        thisSize = GetTitleFontSize(sld)
        If modeSize > 0 And thisSize > 0 And Abs(thisSize - modeSize) > 0.5 Then
            AddIssue sld.slideIndex, "Warning", "Title Font Size", _
                "Title is " & Format(thisSize, "0.#") & "pt; most slides use " & Format(modeSize, "0.#") & "pt", _
                "", GetTitleShapeName(sld)
        End If
    End If

    
    If cfg.T4_TitleFontFace And m_TitleCount > 1 Then
        Dim modeFace As String
        modeFace = GetModeStringForLayout(m_TitleFontNamesArr, m_TitleLayoutNames, m_TitleCount, sld.CustomLayout.name)
        Dim thisFace As String
        thisFace = GetTitleFontName(sld)
        If modeFace <> "" And thisFace <> "" And LCase(thisFace) <> LCase(modeFace) Then
            AddIssue sld.slideIndex, "Warning", "Title Font Face", _
                "Title uses " & thisFace & "; most slides use " & modeFace, _
                "", GetTitleShapeName(sld)
        End If
    End If

    
    If cfg.T5_TitleEndsPeriod Then
        If right(Trim(ttl), 1) = "." Then
            AddIssue sld.slideIndex, "Info", "Title Ends Period", _
                "Title ends with a period: " & Chr(34) & left(ttl, 50) & IIf(Len(ttl) > 50, "...", "") & Chr(34), _
                "T5", GetTitleShapeName(sld)
        End If
    End If

    
    If cfg.T6_DuplicateTitles Then
        Dim lttl As String
        lttl = LCase(Trim(ttl))
        Dim i As Long
        For i = 1 To m_TitleCount
            If m_AllTitleSlideIdx(i) < sld.slideIndex Then
                If m_AllTitles(i) = lttl Then
                    AddIssue sld.slideIndex, "Info", "Duplicate Title", _
                        "Same title as slide " & m_AllTitleSlideIdx(i) & ": " & Chr(34) & left(ttl, 45) & IIf(Len(ttl) > 45, "...", "") & Chr(34), _
                        "", GetTitleShapeName(sld)
                    Exit For
                ElseIf LevenshteinDist(m_AllTitles(i), lttl) <= 2 And Len(lttl) > 5 Then
                    AddIssue sld.slideIndex, "Info", "Near-Duplicate Title", _
                        "Very similar title to slide " & m_AllTitleSlideIdx(i), _
                        "", GetTitleShapeName(sld)
                    Exit For
                End If
            End If
        Next i
    End If

End Sub

Private Sub CheckTextQuality(sld As PowerPoint.Slide, cfg As GraderCheckConfig)

    Dim shp As PowerPoint.shape
    Dim totalBodyWords As Long
    totalBodyWords = 0

    For Each shp In sld.shapes
        On Error Resume Next

        If shp.Type = msoGroup Then GoTo NextShape

        
        If cfg.Q4_EmptyTextBox Then
            If shp.HasTextFrame And shp.Type <> msoPlaceholder Then
                If Trim(shp.TextFrame.textRange.text) = "" Then
                    AddIssue sld.slideIndex, "Warning", "Empty Text Box", _
                        "Shape " & Chr(34) & left(shp.name, 30) & Chr(34) & " has no text", _
                        "Q4", shp.name
                End If
            End If
        End If

        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                Dim fullText As String
                fullText = shp.TextFrame.textRange.text

                
                If Not IsTitlePlaceholder(shp) Then
                    totalBodyWords = totalBodyWords + CountWords(fullText)
                End If

                
                If cfg.Q1_DoubleSpaces Then
                    If InStr(fullText, "  ") > 0 Then
                        AddIssue sld.slideIndex, "Warning", "Double Spaces", _
                            "Double space in " & Chr(34) & left(shp.name, 30) & Chr(34), _
                            "Q1", shp.name
                    End If
                End If

                
                If cfg.Q8_DoublePunct Then
                    If InStr(fullText, "..") > 0 Or InStr(fullText, "!!") > 0 Or _
                       InStr(fullText, "?!") > 0 Or InStr(fullText, "!?") > 0 Then
                        AddIssue sld.slideIndex, "Info", "Double Punctuation", _
                            "Double punctuation in " & Chr(34) & left(shp.name, 30) & Chr(34), _
                            "Q8", shp.name
                    End If
                End If

                
                If cfg.Q3_PlaceholderText Then
                    Dim ltext As String
                    ltext = LCase(fullText)
                    If InStr(ltext, "lorem ipsum") > 0 Or InStr(ltext, "[todo]") > 0 Or _
                       InStr(ltext, "[tbd]") > 0 Or InStr(ltext, "click to add") > 0 Or _
                       InStr(ltext, "click to edit") > 0 Or InStr(ltext, "[text]") > 0 Or _
                       InStr(ltext, "[placeholder]") > 0 Or InStr(ltext, "fixme") > 0 Or _
                       InStr(ltext, "insert text here") > 0 Then
                        AddIssue sld.slideIndex, "Error", "Placeholder Text", _
                            "Unreplaced placeholder in " & Chr(34) & left(shp.name, 30) & Chr(34), _
                            "", shp.name
                    End If
                End If

                
                If cfg.Q2_TrailingSpaces Then
                    Dim p As Long
                    For p = 1 To shp.TextFrame.textRange.Paragraphs.count
                        Dim paraText As String
                        paraText = shp.TextFrame.textRange.Paragraphs(p).text
                        
                        Dim stripped As String
                        stripped = paraText
                        Do While right(stripped, 1) = Chr(13) Or right(stripped, 1) = Chr(10)
                            stripped = left(stripped, Len(stripped) - 1)
                        Loop
                        If Len(stripped) > 0 And (left(stripped, 1) = " " Or right(stripped, 1) = " ") Then
                            AddIssue sld.slideIndex, "Info", "Trailing Space", _
                                "Leading/trailing space in " & Chr(34) & left(shp.name, 30) & Chr(34), _
                                "Q2", shp.name
                            Exit For
                        End If
                    Next p
                End If

                
                If cfg.Q6_BulletPunctuation Then
                    Dim hasPeriod As Boolean, hasNoPeriod As Boolean
                    hasPeriod = False: hasNoPeriod = False
                    For p = 1 To shp.TextFrame.textRange.Paragraphs.count
                        Dim pt As String
                        pt = shp.TextFrame.textRange.Paragraphs(p).text
                        
                        Do While right(pt, 1) = Chr(13) Or right(pt, 1) = Chr(10)
                            pt = left(pt, Len(pt) - 1)
                        Loop
                        If Len(Trim(pt)) > 3 Then
                            If right(Trim(pt), 1) = "." Then
                                hasPeriod = True
                            ElseIf right(Trim(pt), 1) <> ":" Then
                                hasNoPeriod = True
                            End If
                        End If
                    Next p
                    If hasPeriod And hasNoPeriod Then
                        AddIssue sld.slideIndex, "Warning", "Bullet Punctuation", _
                            "Mixed bullet endings (period vs. none) in " & Chr(34) & left(shp.name, 30) & Chr(34), _
                            "", shp.name
                    End If
                End If

                
                If cfg.Q7_Strikethrough Then
                    Dim r As Long
                    For r = 1 To shp.TextFrame2.textRange.Runs.count
                        If shp.TextFrame2.textRange.Runs(r).Font.Strikethrough = msoTrue Then
                            AddIssue sld.slideIndex, "Info", "Strikethrough Text", _
                                "Strikethrough text in " & Chr(34) & left(shp.name, 30) & Chr(34), _
                                "Q7", shp.name
                            Exit For
                        End If
                    Next r
                End If

            End If
        End If

NextShape:
        On Error GoTo 0
    Next shp

    
    If cfg.Q5_ExcessiveWords Then
        If totalBodyWords > cfg.MaxBodyWords Then
            AddIssue sld.slideIndex, "Warning", "Excessive Text", _
                totalBodyWords & " words of body text (max " & cfg.MaxBodyWords & ")"
        End If
    End If

End Sub

Private Sub CheckTypography(sld As PowerPoint.Slide, cfg As GraderCheckConfig)

    Dim shp As PowerPoint.shape
    Dim fontPipe As String
    Dim sizePipe As String
    fontPipe = "|"
    sizePipe = "|"

    For Each shp In sld.shapes
        On Error Resume Next

        If shp.Type = msoGroup Then
            CollectFontsFromShape shp, fontPipe, sizePipe, sld, cfg
            GoTo NextShape2
        End If

        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then

                Dim paraIdx As Long
                For paraIdx = 1 To shp.TextFrame.textRange.Paragraphs.count
                    Dim para As PowerPoint.textRange
                    Set para = shp.TextFrame.textRange.Paragraphs(paraIdx)

                    Dim runIdx As Long
                    For runIdx = 1 To para.Runs.count
                        Dim run As PowerPoint.textRange
                        Set run = para.Runs(runIdx)

                        
                        Dim fn As String
                        fn = run.Font.name
                        If fn <> "" And InStr(fontPipe, "|" & fn & "|") = 0 Then
                            fontPipe = fontPipe & fn & "|"
                        End If

                        
                        Dim fs As Single
                        fs = run.Font.Size
                        If fs > 0 Then
                            Dim fsKey As String
                            fsKey = Format(fs, "0.#")
                            If InStr(sizePipe, "|" & fsKey & "|") = 0 Then
                                sizePipe = sizePipe & fsKey & "|"
                            End If
                        End If

                        
                        If cfg.F3_FontTooSmall Then
                            If fs > 0 And fs < cfg.MinFontSizePt And Not IsTitlePlaceholder(shp) Then
                                AddIssue sld.slideIndex, "Warning", "Font Too Small", _
                                    Format(fs, "0.#") & "pt text in " & Chr(34) & left(shp.name, 30) & Chr(34), _
                                    "", shp.name
                            End If
                        End If

                        
                        If cfg.F4_UnderlineText Then
                            If run.Font.Underline = msoTrue Then
                                AddIssue sld.slideIndex, "Info", "Underlined Text", _
                                    "Underline in " & Chr(34) & left(shp.name, 30) & Chr(34), _
                                    "F4", shp.name
                            End If
                        End If

                    Next runIdx

                    
                    If cfg.F5_AllCapsBody And Not IsTitlePlaceholder(shp) Then
                        Dim pt5 As String
                        pt5 = Trim(para.text)
                        If Len(pt5) > 15 Then
                            Dim cleaned5 As String
                            cleaned5 = pt5
                            
                            Dim k5 As Integer
                            Dim onlyAlpha As String
                            onlyAlpha = ""
                            For k5 = 1 To Len(cleaned5)
                                Dim c5 As String
                                c5 = Mid(cleaned5, k5, 1)
                                If c5 >= "A" And c5 <= "Z" Then onlyAlpha = onlyAlpha & c5
                                If c5 >= "a" And c5 <= "z" Then onlyAlpha = onlyAlpha & c5
                            Next k5
                            If Len(onlyAlpha) > 5 And onlyAlpha = UCase(onlyAlpha) Then
                                AddIssue sld.slideIndex, "Info", "All-Caps Body", _
                                    "Entire paragraph in caps in " & Chr(34) & left(shp.name, 30) & Chr(34), _
                                    "", shp.name
                            End If
                        End If
                    End If

                    
                    If cfg.F7_InvisibleText Then
                        Dim textClr As Long
                        Dim bgClr As Long
                        textClr = para.Font.color.RGB
                        If shp.Fill.visible = msoTrue And shp.Fill.Type = msoFillSolid Then
                            bgClr = shp.Fill.ForeColor.RGB
                            If RGBDist(textClr, bgClr) < 30 Then
                                AddIssue sld.slideIndex, "Error", "Invisible Text", _
                                    "Text color matches background in " & Chr(34) & left(shp.name, 30) & Chr(34), _
                                    "", shp.name
                            End If
                        End If
                    End If

                Next paraIdx

                
                If cfg.F6_BodyColorIncons And Not IsTitlePlaceholder(shp) Then
                    If m_BodyColorCount > 0 Then
                        Dim modeClr As Long
                        modeClr = GetModeBodyColor()
                        Dim bodyClr As Long
                        bodyClr = shp.TextFrame.textRange.Font.color.RGB
                        If modeClr <> 0 And bodyClr <> 0 And RGBDist(bodyClr, modeClr) > 60 Then
                            AddIssue sld.slideIndex, "Info", "Body Color", _
                                "Body text color differs from deck standard in " & Chr(34) & left(shp.name, 30) & Chr(34), _
                                "", shp.name
                        End If
                    End If
                End If

            End If
        End If

        
        If shp.HasTable Then
            Dim tr As Long, tc As Long
            For tr = 1 To shp.table.rows.count
                For tc = 1 To shp.table.Columns.count
                    Dim cellShp As PowerPoint.shape
                    Set cellShp = shp.table.cell(tr, tc).shape
                    Dim cellPipe As String
                    Dim cellSizePipe As String
                    cellPipe = "|": cellSizePipe = "|"
                    CollectFontsFromShape cellShp, fontPipe, sizePipe, sld, cfg
                Next tc
            Next tr
        End If

NextShape2:
        On Error GoTo 0
    Next shp

    
    If cfg.F1_FontFamilyMix Then
        Dim famCount As Long
        famCount = CountPipeParts(fontPipe)
        If famCount > cfg.MaxFontFamilies Then
            AddIssue sld.slideIndex, "Warning", "Font Family Mix", _
                famCount & " font families (max " & cfg.MaxFontFamilies & "): " & First3PipeParts(fontPipe)
        End If
    End If

    
    If cfg.F2_FontSizeMix Then
        Dim sizeCount As Long
        sizeCount = CountPipeParts(sizePipe)
        If sizeCount > cfg.MaxFontSizes Then
            AddIssue sld.slideIndex, "Warning", "Font Size Mix", _
                sizeCount & " distinct font sizes (max " & cfg.MaxFontSizes & ")"
        End If
    End If

End Sub

Private Sub CheckLayout(sld As PowerPoint.Slide, cfg As GraderCheckConfig)

    Dim shp As PowerPoint.shape
    Dim shpList() As PowerPoint.shape
    Dim shpCount As Long
    shpCount = 0
    ReDim shpList(1 To sld.shapes.count + 1)

    Dim slideW As Single
    Dim slideH As Single
    slideW = ActivePresentation.PageSetup.slideWidth
    slideH = ActivePresentation.PageSetup.slideHeight

    Dim contentShapeCount As Long
    contentShapeCount = 0
    Dim hasContentShape As Boolean
    hasContentShape = False

    For Each shp In sld.shapes
        On Error Resume Next

        If shp.Type = msoGroup Then GoTo NextShapeL

        
        shpCount = shpCount + 1
        Set shpList(shpCount) = shp

        
        If shp.Type <> msoPlaceholder Then
            contentShapeCount = contentShapeCount + 1
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then hasContentShape = True
            Else
                hasContentShape = True
            End If
        Else
            If shp.TextFrame.HasText Then hasContentShape = True
        End If

        
        If cfg.L2_OutsideBounds Then
            If shp.left < -1 Or shp.Top < -1 Or _
               shp.left + shp.width > slideW + 1 Or _
               shp.Top + shp.height > slideH + 1 Then
                AddIssue sld.slideIndex, "Error", "Outside Bounds", _
                    Chr(34) & left(shp.name, 30) & Chr(34) & " extends beyond slide edge", _
                    "", shp.name
            End If
        End If

        
        If cfg.L4_InvisibleShape Then
            Dim hasFill As Boolean, hasLine As Boolean, hasTxt As Boolean
            hasFill = (shp.Fill.visible = msoTrue)
            hasLine = (shp.line.visible = msoTrue)
            hasTxt = shp.HasTextFrame And shp.TextFrame.HasText
            If Not hasFill And Not hasLine And Not hasTxt And shp.HasTextFrame Then
                AddIssue sld.slideIndex, "Warning", "Invisible Shape", _
                    Chr(34) & left(shp.name, 30) & Chr(34) & " is invisible (no fill, line, or text)", _
                    "", shp.name
            End If
        End If

        
        If cfg.L5_RotatedTextBox Then
            If shp.HasTextFrame And Abs(shp.rotation) > 0.5 Then
                AddIssue sld.slideIndex, "Info", "Rotated Text Box", _
                    Chr(34) & left(shp.name, 30) & Chr(34) & " is rotated " & Format(shp.rotation, "0.#") & Chr(176), _
                    "L5", shp.name
            End If
        End If

        
        If cfg.L1_TextOverflow Then
            If shp.HasTextFrame Then
                If shp.TextFrame.AutoSize = ppAutoSizeNone And shp.TextFrame.HasText Then
                    Dim textH As Single, interiorH As Single
                    Dim textW As Single, interiorW As Single
                    textH = shp.TextFrame.textRange.BoundHeight
                    interiorH = shp.height - shp.TextFrame.MarginTop - shp.TextFrame.marginBottom
                    textW = shp.TextFrame.textRange.BoundWidth
                    interiorW = shp.width - shp.TextFrame.MarginLeft - shp.TextFrame.MarginRight
                    If textH > interiorH + 2 Then
                        AddIssue sld.slideIndex, "Warning", "Text Overflow", _
                            "Text overflows vertically in " & Chr(34) & left(shp.name, 30) & Chr(34), _
                            "", shp.name
                    ElseIf Not shp.TextFrame.WordWrap And textW > interiorW + 2 Then
                        AddIssue sld.slideIndex, "Warning", "Text Overflow", _
                            "Text overflows horizontally in " & Chr(34) & left(shp.name, 30) & Chr(34), _
                            "", shp.name
                    End If
                End If
            End If
        End If

        
        If cfg.L8_IsolatedFloat And shp.Type <> msoPlaceholder Then
            If Not ShapeSharesEdge(shp, sld) Then
                AddIssue sld.slideIndex, "Info", "Isolated Shape", _
                    Chr(34) & left(shp.name, 30) & Chr(34) & " doesn't align with any other shape", _
                    "", shp.name
            End If
        End If

NextShapeL:
        On Error GoTo 0
    Next shp

    
    If cfg.L3_OverlappingShapes And shpCount >= 2 Then
        Dim i As Long, j As Long
        For i = 1 To shpCount - 1
            For j = i + 1 To shpCount
                On Error Resume Next
                If shpList(i).Type <> msoLine And shpList(i).Connector <> msoTrue And _
                   shpList(j).Type <> msoLine And shpList(j).Connector <> msoTrue Then
                    If BBoxOverlap(shpList(i), shpList(j)) Then
                        AddIssue sld.slideIndex, "Warning", "Overlapping Shapes", _
                            Chr(34) & left(shpList(i).name, 20) & Chr(34) & " overlaps " & Chr(34) & left(shpList(j).name, 20) & Chr(34), _
                            "", shpList(i).name
                        GoTo DoneOverlap
                    End If
                End If
                On Error GoTo 0
            Next j
        Next i
DoneOverlap:
    End If

    
    If cfg.L6_TooManyShapes Then
        If contentShapeCount > cfg.MaxShapesPerSlide Then
            AddIssue sld.slideIndex, "Info", "Too Many Shapes", _
                contentShapeCount & " content shapes (max " & cfg.MaxShapesPerSlide & ")"
        End If
    End If

    
    If cfg.L7_NoContent Then
        If Not hasContentShape Then
            AddIssue sld.slideIndex, "Warning", "No Content", _
                "Slide has no content beyond its title"
        End If
    End If

End Sub

Private Sub CheckColors(sld As PowerPoint.Slide, cfg As GraderCheckConfig)

    Dim shp As PowerPoint.shape
    Dim colorPipe As String
    colorPipe = "|"

    For Each shp In sld.shapes
        On Error Resume Next

        
        If shp.Fill.visible = msoTrue And shp.Fill.Type = msoFillSolid Then
            Dim fc As Long
            fc = shp.Fill.ForeColor.RGB
            If fc <> RGB(255, 255, 255) And fc <> 0 Then
                Dim fcKey As String
                fcKey = CStr(fc)
                If InStr(colorPipe, "|" & fcKey & "|") = 0 Then
                    colorPipe = colorPipe & fcKey & "|"
                End If
            End If
        End If

        
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                Dim tc As Long
                tc = shp.TextFrame.textRange.Font.color.RGB
                If tc <> 0 Then
                    If InStr(colorPipe, "|" & CStr(tc) & "|") = 0 Then
                        colorPipe = colorPipe & CStr(tc) & "|"
                    End If
                End If
            End If
        End If

        
        If cfg.C4_LowContrast Then
            If shp.HasTextFrame And shp.TextFrame.HasText Then
                If shp.Fill.visible = msoTrue And shp.Fill.Type = msoFillSolid Then
                    Dim txtC As Long, bgC As Long
                    txtC = shp.TextFrame.textRange.Font.color.RGB
                    bgC = shp.Fill.ForeColor.RGB
                    If RGBDist(txtC, bgC) < 80 And RGBDist(txtC, bgC) > 0 Then
                        AddIssue sld.slideIndex, "Error", "Low Contrast", _
                            "Low text/background contrast in " & Chr(34) & left(shp.name, 30) & Chr(34), _
                            "", shp.name
                    End If
                End If
            End If
        End If

        On Error GoTo 0
    Next shp

    
    If cfg.C1_TooManyColors Then
        Dim clrCount As Long
        clrCount = CountPipeParts(colorPipe)
        If clrCount > cfg.MaxColorsPerSlide Then
            AddIssue sld.slideIndex, "Warning", "Too Many Colors", _
                clrCount & " distinct colors (max " & cfg.MaxColorsPerSlide & ")"
        End If
    End If

    
    If cfg.C2_FillColorIncons Then
        CheckFillColorConsistency sld
    End If

End Sub

Private Sub CheckFillColorConsistency(sld As PowerPoint.Slide)
    
    Dim shp As PowerPoint.shape
    Dim sizeKey As String
    Dim colorAtSize As String
    colorAtSize = "|"

    For Each shp In sld.shapes
        On Error Resume Next
        If shp.Type <> msoPlaceholder And shp.Type <> msoLine Then
            If shp.Fill.visible = msoTrue And shp.Fill.Type = msoFillSolid Then
                
                Dim rw As Long, rh As Long
                rw = CLng(shp.width / 5) * 5
                rh = CLng(shp.height / 5) * 5
                sizeKey = CStr(rw) & "x" & CStr(rh)
                Dim existingClr As Long
                existingClr = GetColorForSizeKey(colorAtSize, sizeKey)
                If existingClr = -1 Then
             
                    colorAtSize = colorAtSize & sizeKey & "=" & CStr(shp.Fill.ForeColor.RGB) & "|"
                ElseIf existingClr <> shp.Fill.ForeColor.RGB And RGBDist(existingClr, shp.Fill.ForeColor.RGB) > 30 Then
                    AddIssue sld.slideIndex, "Info", "Fill Color Mismatch", _
                        "Same-size shapes have different fills on this slide", _
                        "", shp.name
                    Exit For
                End If
            End If
        End If
        On Error GoTo 0
    Next shp
End Sub

Private Function GetColorForSizeKey(colorAtSize As String, sizeKey As String) As Long
    Dim pattern As String
    pattern = "|" & sizeKey & "="
    Dim pos As Long
    pos = InStr(colorAtSize, pattern)
    If pos = 0 Then
        GetColorForSizeKey = -1
    Else
        Dim startPos As Long
        startPos = pos + Len(pattern)
        Dim endPos As Long
        endPos = InStr(startPos, colorAtSize, "|")
        If endPos > startPos Then
            GetColorForSizeKey = CLng(Mid(colorAtSize, startPos, endPos - startPos))
        Else
            GetColorForSizeKey = -1
        End If
    End If
End Function

Private Sub CheckTablesDeck(sld As PowerPoint.Slide, cfg As GraderCheckConfig)


    If cfg.D3_HiddenSlides Then
        If sld.SlideShowTransition.Hidden = msoTrue Then
            AddIssue sld.slideIndex, "Info", "Hidden Slide", "This slide is hidden from the slideshow", "D3"
        End If
    End If


    If cfg.D4_NotesDraft Then
        On Error Resume Next
        Dim notesText As String
        notesText = LCase(sld.NotesPage.shapes.Placeholders(2).TextFrame.textRange.text)
        If InStr(notesText, "todo") > 0 Or InStr(notesText, "tbc") > 0 Or _
           InStr(notesText, "draft") > 0 Or InStr(notesText, "fixme") > 0 Or _
           InStr(notesText, "placeholder") > 0 Or InStr(notesText, "tbd") > 0 Then
            AddIssue sld.slideIndex, "Warning", "Draft in Notes", _
                "Speaker notes contain draft markers (TODO/TBC/DRAFT)"
        End If
        On Error GoTo 0
    End If

    Dim shp As PowerPoint.shape
    For Each shp In sld.shapes
        On Error Resume Next
        If shp.HasTable Then

         
            If cfg.D1_TableHeader Then
                If shp.table.rows.count >= 2 Then
                    Dim row1Fill As Long, row2Fill As Long
                    row1Fill = shp.table.cell(1, 1).shape.Fill.ForeColor.RGB
                    row2Fill = shp.table.cell(2, 1).shape.Fill.ForeColor.RGB
                    If RGBDist(row1Fill, row2Fill) < 15 Then
                        AddIssue sld.slideIndex, "Warning", "Table Header", _
                            "Table header row not visually distinct from body rows", _
                            "", shp.name
                    End If
                End If
            End If

           
            If cfg.D2_EmptyTableCells Then
                Dim totalCells As Long, filledCells As Long
                totalCells = shp.table.rows.count * shp.table.Columns.count
                filledCells = 0
                Dim tr2 As Long, tc2 As Long
                For tr2 = 1 To shp.table.rows.count
                    For tc2 = 1 To shp.table.Columns.count
                        If shp.table.cell(tr2, tc2).shape.TextFrame.HasText Then
                            filledCells = filledCells + 1
                        End If
                    Next tc2
                Next tr2
             
                If filledCells > totalCells / 2 And filledCells < totalCells Then
                    AddIssue sld.slideIndex, "Info", "Empty Table Cells", _
                        (totalCells - filledCells) & " empty cell(s) in a mostly-filled table", _
                        "", shp.name
                End If
            End If

        End If
        On Error GoTo 0
    Next shp

End Sub


Private Sub AddIssue(slideIndex As Long, sev As String, checkName As String, desc As String, _
                     Optional fixKey As String = "", Optional auxData As String = "")
    m_IssueCount = m_IssueCount + 1
    If m_IssueCount > UBound(m_Issues) Then
        ReDim Preserve m_Issues(1 To m_IssueCount * 2)
    End If
    With m_Issues(m_IssueCount)
        .slideIndex = slideIndex
        .SlideNumber = slideIndex
        .Severity = sev
        .checkName = checkName
        .Description = desc
        .fixKey = fixKey
        .auxData = auxData
        .IsFixed = False
    End With
End Sub

Public Function CanFix(issue As SlideIssue) As Boolean
    CanFix = (Len(issue.fixKey) > 0) And Not issue.IsFixed
End Function

Public Sub FixIssue(issueIdx As Long)
    If issueIdx < 1 Or issueIdx > m_IssueCount Then Exit Sub
    Dim issue As SlideIssue
    issue = m_Issues(issueIdx)
    If Not CanFix(issue) Then Exit Sub

    Select Case issue.fixKey
        Case "T5": Fix_T5 issue
        Case "Q1": Fix_Q1 issue
        Case "Q2": Fix_Q2 issue
        Case "Q4": Fix_Q4 issue
        Case "Q7": Fix_Q7 issue
        Case "Q8": Fix_Q8 issue
        Case "F4": Fix_F4 issue
        Case "L5": Fix_L5 issue
        Case "D3": Fix_D3 issue
    End Select

    m_Issues(issueIdx).IsFixed = True
End Sub


Public Sub FixAllInCategory(catLetter As String)
    Dim i As Long
    For i = 1 To m_IssueCount
        If Not m_Issues(i).IsFixed Then
            If left(m_Issues(i).fixKey, 1) = catLetter Then
                FixIssue i
            End If
        End If
    Next i
End Sub


Private Sub Fix_T5(issue As SlideIssue)
    Dim sld As PowerPoint.Slide
    Set sld = ActivePresentation.Slides(issue.slideIndex)
    Dim shp As PowerPoint.shape
    For Each shp In sld.shapes
        On Error Resume Next
        If shp.Type = msoPlaceholder Then
            If shp.PlaceholderFormat.Type = ppPlaceholderTitle Or _
               shp.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
                Dim t As String
                t = shp.TextFrame.textRange.text
                If right(Trim(t), 1) = "." Then
                    shp.TextFrame.textRange.text = left(Trim(t), Len(Trim(t)) - 1)
                End If
                Exit For
            End If
        End If
        On Error GoTo 0
    Next shp
End Sub


Private Sub Fix_Q1(issue As SlideIssue)
    Dim shp As PowerPoint.shape
    Set shp = FindShapeByName(ActivePresentation.Slides(issue.slideIndex), issue.auxData)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    Dim txt As String
    txt = shp.TextFrame.textRange.text
    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop
    shp.TextFrame.textRange.text = txt
    On Error GoTo 0
End Sub


Private Sub Fix_Q2(issue As SlideIssue)
    Dim shp As PowerPoint.shape
    Set shp = FindShapeByName(ActivePresentation.Slides(issue.slideIndex), issue.auxData)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    Dim p As Long
    For p = 1 To shp.TextFrame.textRange.Paragraphs.count
        Dim para As PowerPoint.textRange
        Set para = shp.TextFrame.textRange.Paragraphs(p)
        Dim cleaned As String
        cleaned = para.text
 
        Do While right(cleaned, 1) = Chr(13) Or right(cleaned, 1) = Chr(10)
            cleaned = left(cleaned, Len(cleaned) - 1)
        Loop
        If cleaned <> Trim(cleaned) Then
            para.text = Trim(cleaned) & Chr(13)
        End If
    Next p
    On Error GoTo 0
End Sub


Private Sub Fix_Q4(issue As SlideIssue)
    Dim shp As PowerPoint.shape
    Set shp = FindShapeByName(ActivePresentation.Slides(issue.slideIndex), issue.auxData)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    shp.Delete
    On Error GoTo 0
End Sub


Private Sub Fix_Q7(issue As SlideIssue)
    Dim shp As PowerPoint.shape
    Set shp = FindShapeByName(ActivePresentation.Slides(issue.slideIndex), issue.auxData)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    Dim r As Long
    For r = 1 To shp.TextFrame2.textRange.Runs.count
        shp.TextFrame2.textRange.Runs(r).Font.Strikethrough = msoFalse
    Next r
    On Error GoTo 0
End Sub


Private Sub Fix_Q8(issue As SlideIssue)
    Dim shp As PowerPoint.shape
    Set shp = FindShapeByName(ActivePresentation.Slides(issue.slideIndex), issue.auxData)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    Dim txt As String
    txt = shp.TextFrame.textRange.text
    txt = Replace(txt, "!?", "?")
    txt = Replace(txt, "?!", "?")
    txt = Replace(txt, "!!", "!")
    Do While InStr(txt, "..") > 0
        txt = Replace(txt, "..", ".")
    Loop
    shp.TextFrame.textRange.text = txt
    On Error GoTo 0
End Sub


Private Sub Fix_F4(issue As SlideIssue)
    Dim shp As PowerPoint.shape
    Set shp = FindShapeByName(ActivePresentation.Slides(issue.slideIndex), issue.auxData)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    Dim r As Long
    For r = 1 To shp.TextFrame2.textRange.Runs.count
        shp.TextFrame2.textRange.Runs(r).Font.UnderlineStyle = msoNoUnderline
    Next r
    On Error GoTo 0
End Sub

Private Sub Fix_L5(issue As SlideIssue)
    Dim shp As PowerPoint.shape
    Set shp = FindShapeByName(ActivePresentation.Slides(issue.slideIndex), issue.auxData)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    shp.rotation = 0
    On Error GoTo 0
End Sub

Private Sub Fix_D3(issue As SlideIssue)
    On Error Resume Next
    ActivePresentation.Slides(issue.slideIndex).SlideShowTransition.Hidden = msoFalse
    On Error GoTo 0
End Sub

Private Function FindShapeByName(sld As PowerPoint.Slide, shapeName As String) As PowerPoint.shape
    Dim shp As PowerPoint.shape
    Set FindShapeByName = Nothing
    If Len(shapeName) = 0 Then Exit Function
    For Each shp In sld.shapes
        If shp.name = shapeName Then
            Set FindShapeByName = shp
            Exit Function
        End If
    Next shp
End Function


Public Sub ClearHighlight()
    If m_HighlightSlideIdx = 0 Then Exit Sub
    On Error Resume Next
    Dim sld As PowerPoint.Slide
    Set sld = ActivePresentation.Slides(m_HighlightSlideIdx)
    If Not sld Is Nothing Then
        Dim shp As PowerPoint.shape
        For Each shp In sld.shapes
            If shp.name = "SlideGraderHighlight" Then
                shp.Delete
                Exit For
            End If
        Next shp
    End If
    On Error GoTo 0
    m_HighlightSlideIdx = 0
End Sub

Public Sub AddHighlight(issue As SlideIssue)
    If Len(issue.auxData) = 0 Then Exit Sub
    On Error Resume Next

    Dim sld As PowerPoint.Slide
    Set sld = ActivePresentation.Slides(issue.slideIndex)
    If sld Is Nothing Then Exit Sub

    Dim target As PowerPoint.shape
    Set target = FindShapeByName(sld, issue.auxData)
    If target Is Nothing Then Exit Sub

    Const PAD As Single = 4
    Dim hl As PowerPoint.shape
    Set hl = sld.shapes.AddShape(msoShapeRectangle, _
        target.left - PAD, target.Top - PAD, _
        target.width + PAD * 2, target.height + PAD * 2)

    With hl
        .name = "SlideGraderHighlight"
        .Fill.visible = msoFalse
        .line.visible = msoTrue
        .line.ForeColor.RGB = RGB(255, 0, 0)
        .line.Weight = 2.5
        .ZOrder msoBringToFront
    End With

    m_HighlightSlideIdx = issue.slideIndex
    On Error GoTo 0
End Sub

Private Function GetSlideTitle(sld As PowerPoint.Slide) As String
    Dim shp As PowerPoint.shape
    GetSlideTitle = ""
    On Error Resume Next
    For Each shp In sld.shapes
        If shp.Type = msoPlaceholder Then
            If shp.PlaceholderFormat.Type = ppPlaceholderTitle Or _
               shp.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
                GetSlideTitle = Trim(shp.TextFrame.textRange.text)
                Exit For
            End If
        End If
    Next shp
    On Error GoTo 0
End Function

Private Function GetTitleShapeName(sld As PowerPoint.Slide) As String
    Dim shp As PowerPoint.shape
    GetTitleShapeName = ""
    On Error Resume Next
    For Each shp In sld.shapes
        If shp.Type = msoPlaceholder Then
            If shp.PlaceholderFormat.Type = ppPlaceholderTitle Or _
               shp.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
                GetTitleShapeName = shp.name
                Exit For
            End If
        End If
    Next shp
    On Error GoTo 0
End Function

Private Function IsTitlePlaceholder(shp As PowerPoint.shape) As Boolean
    IsTitlePlaceholder = False
    On Error Resume Next
    If shp.Type = msoPlaceholder Then
        If shp.PlaceholderFormat.Type = ppPlaceholderTitle Or _
           shp.PlaceholderFormat.Type = ppPlaceholderCenterTitle Or _
           shp.PlaceholderFormat.Type = ppPlaceholderSubtitle Then
            IsTitlePlaceholder = True
        End If
    End If
    On Error GoTo 0
End Function

Private Function TitleHasVerb(titleText As String) As Boolean
    Dim verbs As Variant
    verbs = Array("grow", "grows", "grew", "fall", "falls", "fell", "rise", "rises", "rose", _
                  "decline", "declines", "declined", "increase", "increases", "increased", _
                  "decrease", "decreases", "decreased", "outperform", "outperforms", "outperformed", _
                  "exceed", "exceeds", "exceeded", "drop", "drops", "dropped", "drive", "drives", _
                  "drove", "lead", "leads", "led", "show", "shows", "showed", "reveal", "reveals", _
                  "revealed", "highlight", "highlights", "highlighted", "achieve", "achieves", _
                  "achieved", "reduce", "reduces", "reduced", "improve", "improves", "improved", _
                  "impact", "impacts", "impacted", "accelerate", "accelerates", "confirm", "confirms", _
                  "suggest", "suggests", "indicate", "indicates", "remain", "remains", "slow", "slows", _
                  "represent", "represents", "enable", "enables", "create", "creates", "boost", "boosts", _
                  "cut", "cuts", "save", "saves", "cost", "costs", "generate", "generates", "deliver", _
                  "delivers", "expand", "expands", "contract", "contracts", "miss", "misses", "beat", _
                  "beats", "gain", "gains", "lose", "loses", "add", "adds")
    Dim cleaned As String
    cleaned = LCase(Replace(Replace(Replace(titleText, ",", " "), ".", " "), ":", " "))
    Dim words() As String
    words = Split(cleaned, " ")
    Dim v As Variant
    Dim w As String
    Dim i As Integer
    For i = LBound(words) To UBound(words)
        w = Trim(words(i))
        For Each v In verbs
            If w = CStr(v) Then
                TitleHasVerb = True
                Exit Function
            End If
        Next v
    Next i
    TitleHasVerb = False
End Function

Private Function CountWords(s As String) As Integer
    If Trim(s) = "" Then CountWords = 0: Exit Function
    Dim cleaned As String
    cleaned = Trim(s)
    Do While InStr(cleaned, "  ") > 0
        cleaned = Replace(cleaned, "  ", " ")
    Loop
    CountWords = UBound(Split(cleaned, " ")) - LBound(Split(cleaned, " ")) + 1
End Function

Private Sub CollectFontsFromShape(shp As PowerPoint.shape, ByRef fontPipe As String, ByRef sizePipe As String, sld As PowerPoint.Slide, cfg As GraderCheckConfig)
    On Error Resume Next
    If shp.Type = msoGroup Then
        Dim child As PowerPoint.shape
        For Each child In shp.GroupItems
            CollectFontsFromShape child, fontPipe, sizePipe, sld, cfg
        Next child
        Exit Sub
    End If
    If shp.HasTextFrame And shp.TextFrame.HasText Then
        Dim p As Long, r As Long
        For p = 1 To shp.TextFrame.textRange.Paragraphs.count
            For r = 1 To shp.TextFrame.textRange.Paragraphs(p).Runs.count
                Dim fn As String
                Dim fs As Single
                fn = shp.TextFrame.textRange.Paragraphs(p).Runs(r).Font.name
                fs = shp.TextFrame.textRange.Paragraphs(p).Runs(r).Font.Size
                If fn <> "" And InStr(fontPipe, "|" & fn & "|") = 0 Then
                    fontPipe = fontPipe & fn & "|"
                End If
                If fs > 0 Then
                    Dim fsk As String
                    fsk = Format(fs, "0.#")
                    If InStr(sizePipe, "|" & fsk & "|") = 0 Then sizePipe = sizePipe & fsk & "|"
                End If
            Next r
        Next p
    End If
    On Error GoTo 0
End Sub

Private Function CountPipeParts(pipe As String) As Long

    If pipe = "|" Or pipe = "" Then CountPipeParts = 0: Exit Function
    Dim trimmed As String
    trimmed = Mid(pipe, 2, Len(pipe) - 2)
    If trimmed = "" Then CountPipeParts = 0: Exit Function
    CountPipeParts = UBound(Split(trimmed, "|")) - LBound(Split(trimmed, "|")) + 1
End Function

Private Function First3PipeParts(pipe As String) As String
    If pipe = "|" Or pipe = "" Then First3PipeParts = "": Exit Function
    Dim trimmed As String
    trimmed = Mid(pipe, 2, Len(pipe) - 2)
    Dim parts() As String
    parts = Split(trimmed, "|")
    Dim result As String
    Dim i As Long
    For i = LBound(parts) To IIf(LBound(parts) + 2 < UBound(parts), LBound(parts) + 2, UBound(parts))
        If result <> "" Then result = result & ", "
        result = result & parts(i)
    Next i
    If UBound(parts) > LBound(parts) + 2 Then result = result & "..."
    First3PipeParts = result
End Function

Private Function RGBDist(c1 As Long, c2 As Long) As Double

    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    r1 = c1 Mod 256
    g1 = (c1 \ 256) Mod 256
    b1 = (c1 \ 65536) Mod 256
    r2 = c2 Mod 256
    g2 = (c2 \ 256) Mod 256
    b2 = (c2 \ 65536) Mod 256
    RGBDist = Sqr((r1 - r2) ^ 2 + (g1 - g2) ^ 2 + (b1 - b2) ^ 2)
End Function

Private Function LevenshteinDist(s1 As String, s2 As String) As Integer

    Dim m As Integer, n As Integer
    m = Len(s1): n = Len(s2)
    If m = 0 Then LevenshteinDist = n: Exit Function
    If n = 0 Then LevenshteinDist = m: Exit Function
    Dim d() As Integer
    ReDim d(0 To m, 0 To n)
    Dim i As Integer, j As Integer
    Dim cost As Integer, a As Integer, b As Integer, c As Integer
    For i = 0 To m: d(i, 0) = i: Next i
    For j = 0 To n: d(0, j) = j: Next j
    For i = 1 To m
        For j = 1 To n
            cost = IIf(Mid(s1, i, 1) = Mid(s2, j, 1), 0, 1)
            a = d(i - 1, j) + 1
            b = d(i, j - 1) + 1
            c = d(i - 1, j - 1) + cost
            d(i, j) = IIf(a < b, IIf(a < c, a, c), IIf(b < c, b, c))
        Next j
    Next i
    LevenshteinDist = d(m, n)
End Function

Private Function BBoxOverlap(a As PowerPoint.shape, b As PowerPoint.shape) As Boolean
    BBoxOverlap = Not (a.left + a.width <= b.left Or b.left + b.width <= a.left Or _
                       a.Top + a.height <= b.Top Or b.Top + b.height <= a.Top)
End Function

Private Function ShapeSharesEdge(shp As PowerPoint.shape, sld As PowerPoint.Slide) As Boolean

    Const TOLERANCE As Single = 3
    Dim other As PowerPoint.shape
    ShapeSharesEdge = False
    For Each other In sld.shapes
        On Error Resume Next
        If other.name <> shp.name Then
            If Abs(shp.left - other.left) < TOLERANCE Or _
               Abs((shp.left + shp.width) - (other.left + other.width)) < TOLERANCE Or _
               Abs(shp.Top - other.Top) < TOLERANCE Or _
               Abs((shp.Top + shp.height) - (other.Top + other.height)) < TOLERANCE Then
                ShapeSharesEdge = True
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next other
End Function

Private Function GetTitleFontSize(sld As PowerPoint.Slide) As Single
    Dim shp As PowerPoint.shape
    GetTitleFontSize = 0
    On Error Resume Next
    For Each shp In sld.shapes
        If shp.Type = msoPlaceholder Then
            If shp.PlaceholderFormat.Type = ppPlaceholderTitle Or _
               shp.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
                GetTitleFontSize = shp.TextFrame.textRange.Font.Size
                Exit For
            End If
        End If
    Next shp
    On Error GoTo 0
End Function

Private Function GetTitleFontName(sld As PowerPoint.Slide) As String
    Dim shp As PowerPoint.shape
    GetTitleFontName = ""
    On Error Resume Next
    For Each shp In sld.shapes
        If shp.Type = msoPlaceholder Then
            If shp.PlaceholderFormat.Type = ppPlaceholderTitle Or _
               shp.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
                GetTitleFontName = shp.TextFrame.textRange.Font.name
                Exit For
            End If
        End If
    Next shp
    On Error GoTo 0
End Function


Private Function GetModeSingleForLayout(arr() As Single, layouts() As String, n As Long, layoutName As String) As Single
    If n = 0 Then GetModeSingleForLayout = 0: Exit Function

    Dim vals() As Single
    Dim cnt As Long
    ReDim vals(1 To n)
    cnt = 0
    Dim i As Long
    For i = 1 To n
        If layouts(i) = layoutName Then
            cnt = cnt + 1
            vals(cnt) = arr(i)
        End If
    Next i
    If cnt = 0 Then

        cnt = n
        For i = 1 To n: vals(i) = arr(i): Next i
    End If
   
    Dim bestVal As Single, bestCnt As Long
    bestCnt = 0
    For i = 1 To cnt
        Dim thisCnt As Long
        thisCnt = 0
        Dim j As Long
        For j = 1 To cnt
            If Abs(vals(i) - vals(j)) < 0.5 Then thisCnt = thisCnt + 1
        Next j
        If thisCnt > bestCnt Then bestCnt = thisCnt: bestVal = vals(i)
    Next i
    GetModeSingleForLayout = bestVal
End Function


Private Function GetModeStringForLayout(arr() As String, layouts() As String, n As Long, layoutName As String) As String
    If n = 0 Then GetModeStringForLayout = "": Exit Function
    Dim vals() As String
    Dim cnt As Long
    ReDim vals(1 To n)
    cnt = 0
    Dim i As Long
    For i = 1 To n
        If layouts(i) = layoutName Then
            cnt = cnt + 1
            vals(cnt) = arr(i)
        End If
    Next i
    If cnt = 0 Then
        cnt = n
        For i = 1 To n: vals(i) = arr(i): Next i
    End If
    Dim bestVal As String, bestCnt As Long
    bestCnt = 0
    For i = 1 To cnt
        Dim thisCnt As Long
        thisCnt = 0
        Dim j As Long
        For j = 1 To cnt
            If LCase(vals(i)) = LCase(vals(j)) Then thisCnt = thisCnt + 1
        Next j
        If thisCnt > bestCnt Then bestCnt = thisCnt: bestVal = vals(i)
    Next i
    GetModeStringForLayout = bestVal
End Function

Private Function GetModeBodyColor() As Long
    If m_BodyColorCount = 0 Then GetModeBodyColor = 0: Exit Function
    Dim bestClr As Long, bestCnt As Long
    bestCnt = 0
    Dim i As Long, j As Long
    For i = 1 To m_BodyColorCount
        Dim thisCnt As Long
        thisCnt = 0
        For j = 1 To m_BodyColorCount
            If m_BodyColorsArr(i) = m_BodyColorsArr(j) Then thisCnt = thisCnt + 1
        Next j
        If thisCnt > bestCnt Then bestCnt = thisCnt: bestClr = m_BodyColorsArr(i)
    Next i
    GetModeBodyColor = bestClr
End Function


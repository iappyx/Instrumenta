VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SlideGraderForm 
   Caption         =   "Slide grader"
   ClientHeight    =   9330.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12510
   OleObjectBlob   =   "SlideGraderForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SlideGraderForm"
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


Private Type CheckDef
    key         As String
    Category    As String
    caption     As String
    defaultOn   As Boolean
End Type

Private m_CheckDefs()      As CheckDef
Private m_CheckCount       As Integer
Private m_CurrentCfg       As GraderCheckConfig
Private m_CatHandlers      As Collection
Private m_UpdatingFilter   As Boolean


Private Sub UserForm_Initialize()

    BuildCheckDefs

    cmdFixIssue.enabled = False
    cmdFixCategory.enabled = False

    With lstIssues
        .ColumnCount = 4
        .ColumnWidths = "20;40;120;0"
    End With


    With cboSort
        .AddItem "By Slide"
        .AddItem "By Severity"
        .AddItem "By Check Type"
        .ListIndex = 0
    End With


    With cboFilter
        .AddItem "All"
        .ListIndex = 0
    End With

    BuildCheckboxes


    lblStatus.caption = "Press Scan to analyse the presentation."
    lblIssue.caption = ""

End Sub


Private Sub BuildCheckDefs()

    m_CheckCount = 37
    ReDim m_CheckDefs(1 To m_CheckCount)

    Dim i As Integer: i = 0

    i = i + 1: m_CheckDefs(i) = MakeDef("T1", "Titles", "Missing or empty title", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("T2", "Titles", "Action title (topic, not insight)", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("T3", "Titles", "Title font size inconsistent across slides", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("T4", "Titles", "Title font face inconsistent across slides", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("T5", "Titles", "Title ends with a period", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("T6", "Titles", "Duplicate or near-duplicate titles", False)

    i = i + 1: m_CheckDefs(i) = MakeDef("Q1", "Text", "Double spaces in text", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("Q2", "Text", "Leading / trailing spaces in paragraph", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("Q3", "Text", "Placeholder text not replaced", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("Q4", "Text", "Empty visible text box", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("Q5", "Text", "Excessive word count on slide (>150)", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("Q6", "Text", "Inconsistent bullet ending punctuation", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("Q7", "Text", "Strikethrough text present", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("Q8", "Text", "Double punctuation (.., !!, ?!)", False)

    i = i + 1: m_CheckDefs(i) = MakeDef("F1", "Typography", "Too many font families on one slide (>2)", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("F2", "Typography", "Too many distinct font sizes on one slide (>4)", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("F3", "Typography", "Body text font size too small (<10 pt)", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("F4", "Typography", "Underlined text present", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("F5", "Typography", "All-caps body paragraph", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("F6", "Typography", "Body text color inconsistent vs. deck", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("F7", "Typography", "Text invisible against shape background", False)

    i = i + 1: m_CheckDefs(i) = MakeDef("L1", "Layout", "Text overflow (text exceeds shape bounds)", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("L2", "Layout", "Object outside slide bounds", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("L3", "Layout", "Overlapping shapes", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("L4", "Layout", "Invisible shape (no fill, line, or text)", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("L5", "Layout", "Rotated text box", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("L6", "Layout", "Too many content shapes on one slide (>20)", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("L7", "Layout", "Slide has no content shapes", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("L8", "Layout", "Isolated shape (not aligned with any other)", False)

    i = i + 1: m_CheckDefs(i) = MakeDef("C1", "Colors", "Too many distinct colors on slide (>5)", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("C2", "Colors", "Same-size shapes have different fill colors", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("C3", "Colors", "Accent color inconsistent across slides", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("C4", "Colors", "Low contrast: text vs. background", False)

    i = i + 1: m_CheckDefs(i) = MakeDef("D1", "Tables & Deck", "Table header not visually distinct from body", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("D2", "Tables & Deck", "Empty cells in mostly-filled table", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("D3", "Tables & Deck", "Hidden slides present", False)
    i = i + 1: m_CheckDefs(i) = MakeDef("D4", "Tables & Deck", "Speaker notes contain draft markers", False)

End Sub

Private Function MakeDef(key As String, cat As String, caption As String, defaultOn As Boolean) As CheckDef
    Dim d As CheckDef
    d.key = key
    d.Category = cat
    d.caption = caption
    d.defaultOn = defaultOn
    MakeDef = d
End Function

Private Sub BuildCheckboxes()

    Dim ctrl As MSForms.control
    For Each ctrl In frmChecks.Controls
        frmChecks.Controls.Remove ctrl.name
    Next ctrl

    Set m_CatHandlers = New Collection

    Const ROW_H   As Single = 15
    Const CAT_H   As Single = 18
    Const LEFT_   As Single = 6
    Const CAT_LEFT As Single = 2
    Const CTRL_W  As Single = 340

    Dim currentY As Single
    currentY = 4

    Dim selAll As MSForms.CheckBox
    Set selAll = frmChecks.Controls.Add("Forms.CheckBox.1", "chkSelectAll")
    With selAll
        .caption = "Select All"
        .value = False
        .Font.Bold = True
        .Font.Size = 9
        .left = CAT_LEFT
        .Top = currentY
        .width = CTRL_W
        .height = CAT_H
    End With
    currentY = currentY + CAT_H + 4

    Dim saHandler As SlideGraderCategoryChk
    Set saHandler = New SlideGraderCategoryChk
    Set saHandler.chk = selAll
    saHandler.CategoryPrefix = ""
    Set saHandler.TargetForm = Me
    m_CatHandlers.Add saHandler

    Dim lastCat As String
    lastCat = ""

    Dim i As Integer
    For i = 1 To m_CheckCount


        If m_CheckDefs(i).Category <> lastCat Then
            If currentY > 4 Then currentY = currentY + 4

            Dim catPrefix As String
            catPrefix = left(m_CheckDefs(i).key, 1)

            Dim catChk As MSForms.CheckBox
            Set catChk = frmChecks.Controls.Add("Forms.CheckBox.1", "cat_" & catPrefix)
            With catChk
                .caption = m_CheckDefs(i).Category
                .value = False
                .Font.Bold = True
                .Font.Size = 8
                .left = CAT_LEFT
                .Top = currentY
                .width = CTRL_W
                .height = CAT_H
            End With
            currentY = currentY + CAT_H
            lastCat = m_CheckDefs(i).Category

            Dim handler As SlideGraderCategoryChk
            Set handler = New SlideGraderCategoryChk
            Set handler.chk = catChk
            handler.CategoryPrefix = catPrefix
            Set handler.TargetForm = Me
            m_CatHandlers.Add handler

        End If

        Dim chk As MSForms.CheckBox
        Set chk = frmChecks.Controls.Add("Forms.CheckBox.1", "chk_" & m_CheckDefs(i).key)
        With chk
            .caption = m_CheckDefs(i).caption
            .value = m_CheckDefs(i).defaultOn
            .left = LEFT_
            .Top = currentY
            .width = CTRL_W
            .height = ROW_H
            .Font.Size = 8
        End With
        currentY = currentY + ROW_H

    Next i

    frmChecks.ScrollBars = fmScrollBarsVertical
    frmChecks.ScrollHeight = currentY + 10
    frmChecks.KeepScrollBarsVisible = fmScrollBarsNone

End Sub

Public Sub ToggleCategoryChildren(prefix As String, checked As Boolean)
    Dim ctrl As MSForms.control
    For Each ctrl In frmChecks.Controls
        If left(ctrl.name, 4) = "chk_" And Mid(ctrl.name, 5, 1) = prefix Then
            ctrl.value = checked
        End If
    Next ctrl
End Sub

Public Sub ToggleAllChildren(checked As Boolean)
    Dim ctrl As MSForms.control
    For Each ctrl In frmChecks.Controls
        If left(ctrl.name, 4) = "cat_" Or left(ctrl.name, 4) = "chk_" Then
            ctrl.value = checked
        End If
    Next ctrl
End Sub


Private Function BuildConfig() As GraderCheckConfig

    Dim cfg As GraderCheckConfig
    cfg = ModuleSlideGrader.GetDefaultConfig()

    cfg.T1_MissingTitle = GetChkValue("chk_T1")
    cfg.T2_ActionTitle = GetChkValue("chk_T2")
    cfg.T3_TitleFontSize = GetChkValue("chk_T3")
    cfg.T4_TitleFontFace = GetChkValue("chk_T4")
    cfg.T5_TitleEndsPeriod = GetChkValue("chk_T5")
    cfg.T6_DuplicateTitles = GetChkValue("chk_T6")
    cfg.Q1_DoubleSpaces = GetChkValue("chk_Q1")
    cfg.Q2_TrailingSpaces = GetChkValue("chk_Q2")
    cfg.Q3_PlaceholderText = GetChkValue("chk_Q3")
    cfg.Q4_EmptyTextBox = GetChkValue("chk_Q4")
    cfg.Q5_ExcessiveWords = GetChkValue("chk_Q5")
    cfg.Q6_BulletPunctuation = GetChkValue("chk_Q6")
    cfg.Q7_Strikethrough = GetChkValue("chk_Q7")
    cfg.Q8_DoublePunct = GetChkValue("chk_Q8")
    cfg.F1_FontFamilyMix = GetChkValue("chk_F1")
    cfg.F2_FontSizeMix = GetChkValue("chk_F2")
    cfg.F3_FontTooSmall = GetChkValue("chk_F3")
    cfg.F4_UnderlineText = GetChkValue("chk_F4")
    cfg.F5_AllCapsBody = GetChkValue("chk_F5")
    cfg.F6_BodyColorIncons = GetChkValue("chk_F6")
    cfg.F7_InvisibleText = GetChkValue("chk_F7")
    cfg.L1_TextOverflow = GetChkValue("chk_L1")
    cfg.L2_OutsideBounds = GetChkValue("chk_L2")
    cfg.L3_OverlappingShapes = GetChkValue("chk_L3")
    cfg.L4_InvisibleShape = GetChkValue("chk_L4")
    cfg.L5_RotatedTextBox = GetChkValue("chk_L5")
    cfg.L6_TooManyShapes = GetChkValue("chk_L6")
    cfg.L7_NoContent = GetChkValue("chk_L7")
    cfg.L8_IsolatedFloat = GetChkValue("chk_L8")
    cfg.C1_TooManyColors = GetChkValue("chk_C1")
    cfg.C2_FillColorIncons = GetChkValue("chk_C2")
    cfg.C3_AccentIncons = GetChkValue("chk_C3")
    cfg.C4_LowContrast = GetChkValue("chk_C4")
    cfg.D1_TableHeader = GetChkValue("chk_D1")
    cfg.D2_EmptyTableCells = GetChkValue("chk_D2")
    cfg.D3_HiddenSlides = GetChkValue("chk_D3")
    cfg.D4_NotesDraft = GetChkValue("chk_D4")

    BuildConfig = cfg
End Function

Private Function GetChkValue(ctrlName As String) As Boolean
    On Error Resume Next
    GetChkValue = CBool(frmChecks.Controls(ctrlName).value)
    On Error GoTo 0
End Function

Private Sub cmdScan_Click()

    lblStatus.caption = "Scanning..."
    lstIssues.Clear
    DoEvents

    Dim cfg As GraderCheckConfig
    cfg = BuildConfig()

    ModuleSlideGrader.RunSlideGrader cfg
    m_CurrentCfg = cfg

    PopulateList

End Sub


Private Sub PopulateList()

    lblIssue.caption = ""
    lstIssues.Clear
    cmdFixIssue.enabled = False
    cmdFixCategory.enabled = False

    Dim total As Long
    total = ModuleSlideGrader.GetIssueCount()

    Dim slideCount As Long
    If Application.Presentations.count > 0 Then slideCount = ActivePresentation.Slides.count

    If total = 0 Then
        lblStatus.caption = "No issues found across " & slideCount & " slides."
        Exit Sub
    End If


    Dim prevFilter As String
    prevFilter = cboFilter.text
    m_UpdatingFilter = True
    cboFilter.Clear
    cboFilter.AddItem "All"
    Dim seenNames As String
    seenNames = "|"
    Dim j As Long
    For j = 1 To total
        Dim fIss As SlideIssue
        fIss = ModuleSlideGrader.GetIssue(j)
        If Not fIss.IsFixed Then
            If InStr(seenNames, "|" & fIss.checkName & "|") = 0 Then
                cboFilter.AddItem fIss.checkName
                seenNames = seenNames & fIss.checkName & "|"
            End If
        End If
    Next j
    cboFilter.ListIndex = 0
    If Len(prevFilter) > 0 And prevFilter <> "All" Then
        Dim fi As Integer
        For fi = 1 To cboFilter.ListCount - 1
            If cboFilter.List(fi) = prevFilter Then
                cboFilter.ListIndex = fi
                Exit For
            End If
        Next fi
    End If
    m_UpdatingFilter = False

    Dim sortedIdx() As Long
    ReDim sortedIdx(1 To total)
    Dim i As Long
    For i = 1 To total: sortedIdx(i) = i: Next i
    SortIssueIndices sortedIdx, total, cboSort.ListIndex

    Dim activeFilter As String
    activeFilter = cboFilter.text

    Dim idxStr As String
    Dim visibleCount As Long
    visibleCount = 0

    For i = 1 To total
        Dim issue As SlideIssue
        issue = ModuleSlideGrader.GetIssue(sortedIdx(i))
        If issue.IsFixed Then GoTo NextIssue
        If activeFilter <> "All" And issue.checkName <> activeFilter Then GoTo NextIssue

        lstIssues.AddItem CStr(issue.SlideNumber)
        lstIssues.List(lstIssues.ListCount - 1, 1) = issue.Severity
        lstIssues.List(lstIssues.ListCount - 1, 2) = issue.checkName
        lstIssues.List(lstIssues.ListCount - 1, 3) = issue.Description
        idxStr = idxStr & CStr(sortedIdx(i)) & "|"
        visibleCount = visibleCount + 1

NextIssue:
    Next i

    lstIssues.Tag = idxStr


    If visibleCount = 0 Then
        lblStatus.caption = "No issues found across " & slideCount & " slides."
    ElseIf visibleCount = 1 Then
        lblStatus.caption = "1 issue found across " & slideCount & " slides."
    Else
        lblStatus.caption = visibleCount & " issues found across " & slideCount & " slides."
    End If

End Sub


Private Sub SortIssueIndices(idx() As Long, n As Long, sortMode As Integer)

    Dim i As Long, j As Long
    Dim tmp As Long
    For i = 2 To n
        tmp = idx(i)
        j = i - 1
        Do While j >= 1
            If CompareIssues(idx(j), tmp, sortMode) > 0 Then
                idx(j + 1) = idx(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        idx(j + 1) = tmp
    Next i
End Sub

Private Function CompareIssues(a As Long, b As Long, sortMode As Integer) As Integer

    Dim ia As SlideIssue, ib As SlideIssue
    ia = ModuleSlideGrader.GetIssue(a)
    ib = ModuleSlideGrader.GetIssue(b)

    Select Case sortMode
        Case 1
            Dim sevA As Integer, sevB As Integer
            sevA = SeverityRank(ia.Severity)
            sevB = SeverityRank(ib.Severity)
            If sevA <> sevB Then
                CompareIssues = IIf(sevA < sevB, -1, 1)
            Else
                CompareIssues = IIf(ia.slideIndex < ib.slideIndex, -1, IIf(ia.slideIndex > ib.slideIndex, 1, 0))
            End If
        Case 2
            If LCase(ia.checkName) < LCase(ib.checkName) Then
                CompareIssues = -1
            ElseIf LCase(ia.checkName) > LCase(ib.checkName) Then
                CompareIssues = 1
            Else
                CompareIssues = IIf(ia.slideIndex < ib.slideIndex, -1, IIf(ia.slideIndex > ib.slideIndex, 1, 0))
            End If
        Case Else
            CompareIssues = IIf(ia.slideIndex < ib.slideIndex, -1, IIf(ia.slideIndex > ib.slideIndex, 1, 0))
    End Select
End Function

Private Function SeverityRank(sev As String) As Integer
    Select Case LCase(sev)
        Case "error":   SeverityRank = 1
        Case "warning": SeverityRank = 2
        Case Else:      SeverityRank = 3
    End Select
End Function


Private Sub lstIssues_Click()

    Dim selectedRow As Long
    selectedRow = lstIssues.ListIndex

    If selectedRow < 0 Then Exit Sub
    If ModuleSlideGrader.GetIssueCount() = 0 Then Exit Sub

 
    Dim idxStr As String
    idxStr = lstIssues.Tag
    If Len(idxStr) = 0 Then Exit Sub

    Dim parts() As String
    parts = Split(left(idxStr, Len(idxStr) - 1), "|")

    If selectedRow > UBound(parts) Then Exit Sub

    Dim originalIdx As Long
    originalIdx = CLng(parts(selectedRow))

    Dim issue As SlideIssue
    issue = ModuleSlideGrader.GetIssue(originalIdx)

    ModuleSlideGrader.ClearHighlight

    On Error Resume Next
    ActiveWindow.View.GotoSlide issue.slideIndex
    On Error GoTo 0

    ModuleSlideGrader.AddHighlight issue
    cmdFixIssue.enabled = ModuleSlideGrader.CanFix(issue)
    cmdFixCategory.enabled = ModuleSlideGrader.CanFix(issue)
    lblIssue.caption = lstIssues.List(selectedRow, 3)

End Sub

Private Sub cboSort_Change()
    If ModuleSlideGrader.GetIssueCount() > 0 Then
        PopulateList
    End If
End Sub


Private Sub cboFilter_Change()
    If m_UpdatingFilter Then Exit Sub
    If ModuleSlideGrader.GetIssueCount() > 0 Then PopulateList
End Sub


Private Sub cmdFixCategory_Click()

    Dim selectedRow As Long
    selectedRow = lstIssues.ListIndex
    If selectedRow < 0 Then Exit Sub

    Dim idxStr As String
    idxStr = lstIssues.Tag
    If Len(idxStr) = 0 Then Exit Sub

    Dim parts() As String
    parts = Split(left(idxStr, Len(idxStr) - 1), "|")
    If selectedRow > UBound(parts) Then Exit Sub

    Dim originalIdx As Long
    originalIdx = CLng(parts(selectedRow))

    Dim issue As SlideIssue
    issue = ModuleSlideGrader.GetIssue(originalIdx)
    If Not ModuleSlideGrader.CanFix(issue) Then Exit Sub

    Dim catLetter As String
    catLetter = left(issue.fixKey, 1)

    ModuleSlideGrader.ClearHighlight
    ModuleSlideGrader.FixAllInCategory catLetter

    cmdFixIssue.enabled = False
    cmdFixCategory.enabled = False
    PopulateList

End Sub


Private Sub cmdFixIssue_Click()

    Dim selectedRow As Long
    selectedRow = lstIssues.ListIndex
    If selectedRow < 0 Then Exit Sub

    Dim idxStr As String
    idxStr = lstIssues.Tag
    If Len(idxStr) = 0 Then Exit Sub

    Dim parts() As String
    parts = Split(left(idxStr, Len(idxStr) - 1), "|")
    If selectedRow > UBound(parts) Then Exit Sub

    Dim originalIdx As Long
    originalIdx = CLng(parts(selectedRow))

    ModuleSlideGrader.ClearHighlight
    ModuleSlideGrader.FixIssue originalIdx

    cmdFixIssue.enabled = False
    PopulateList

End Sub

Private Sub cmdClose_Click()
    ModuleSlideGrader.ClearHighlight
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ModuleSlideGrader.ClearHighlight
End Sub



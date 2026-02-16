VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PyramidForm 
   Caption         =   "Pyramid builder"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10665
   OleObjectBlob   =   "PyramidForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PyramidForm"
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

Private Type PyramidNode
    text As String
    depth As Long
    parentIndex As Long
End Type

Private nodes() As PyramidNode
Private nodeCount As Long
Private selectedIndex As Long

Private Sub btnClear_Click()
RemoveAllInstrumentaPyramidTags
Unload Me
End Sub

Private Sub UserForm_Initialize()
    selectedIndex = -1
    nodeCount = 0
        
    optCurrentPres.Value = True
    
    Call LoadPyramidFromTags
    
    If nodeCount = 0 Then
        Call AddNode("Answer", 0, -1)
        Call AddNode("Key Line 1", 1, 0)
        Call AddNode("Key Line 2", 1, 0)
    End If
    
    Call RefreshList
End Sub

Private Sub lstPyramid_Click()
    If lstPyramid.ListIndex >= 0 Then
        selectedIndex = lstPyramid.ListIndex
        txtNodeText.text = nodes(selectedIndex).text
    End If
End Sub

Private Sub btnAddChild_Click()
    Dim newText As String
    
    If selectedIndex < 0 Then
        MsgBox "Please select a parent node first.", vbExclamation
        Exit Sub
    End If
    
    newText = Trim(txtNodeText.text)
    If newText = "" Then
        MsgBox "Please enter text for the new node.", vbExclamation
        Exit Sub
    End If
    
    Call AddNode(newText, nodes(selectedIndex).depth + 1, selectedIndex)
    txtNodeText.text = ""
    Call RefreshList
End Sub

Private Sub btnPromote_Click()
    
    If selectedIndex <= 0 Then Exit Sub
    If selectedIndex = 1 Then Exit Sub
    
    If nodes(selectedIndex).depth <= 1 Then
        MsgBox "Cannot promote this node further.", vbExclamation
        Exit Sub
    End If
    
    If HasChildren(selectedIndex) Then
        MsgBox "Cannot promote a node with children. Promote children first or delete them.", vbExclamation
        Exit Sub
    End If
    
    Dim oldParentIdx As Long
    Dim newParentIdx As Long
    
    oldParentIdx = nodes(selectedIndex).parentIndex
    newParentIdx = nodes(oldParentIdx).parentIndex
    
    nodes(selectedIndex).depth = nodes(selectedIndex).depth - 1
    nodes(selectedIndex).parentIndex = newParentIdx
    
    Call RefreshList
End Sub

Private Sub btnDemote_Click()
    
    If selectedIndex <= 0 Then Exit Sub
    
    Dim prevSiblingIdx As Long
    prevSiblingIdx = FindPreviousSibling(selectedIndex)
    
    If prevSiblingIdx = -1 Then
        MsgBox "Cannot demote. No previous sibling to become parent.", vbExclamation
        Exit Sub
    End If
    
    If HasChildren(selectedIndex) Then
        MsgBox "Cannot demote a node with children. Move children first or delete them.", vbExclamation
        Exit Sub
    End If
    
    nodes(selectedIndex).depth = nodes(selectedIndex).depth + 1
    nodes(selectedIndex).parentIndex = prevSiblingIdx
    
    Call RefreshList
End Sub

Private Function HasChildren(idx As Long) As Long
    Dim i As Long
    
    For i = 0 To nodeCount - 1
        If nodes(i).parentIndex = idx Then
            HasChildren = True
            Exit Function
        End If
    Next i
    
    HasChildren = False
End Function

Private Sub btnEdit_Click()
    Dim newText As String
    
    If selectedIndex < 0 Then
        MsgBox "Please select a node to edit.", vbExclamation
        Exit Sub
    End If
    
    newText = Trim(txtNodeText.text)
    If newText = "" Then
        MsgBox "Node text cannot be empty.", vbExclamation
        Exit Sub
    End If
    
    nodes(selectedIndex).text = newText
    Call RefreshList
End Sub

Private Sub btnDelete_Click()
    Dim i As Long
    Dim newNodes() As PyramidNode
    Dim newCount As Long
    
    If selectedIndex < 0 Then
        MsgBox "Please select a node to delete.", vbExclamation
        Exit Sub
    End If
    
    If selectedIndex = 0 Then
        MsgBox "Cannot delete the Answer node.", vbExclamation
        Exit Sub
    End If
    
    For i = 0 To nodeCount - 1
        If nodes(i).parentIndex = selectedIndex Then
            MsgBox "Cannot delete a node with children. Delete children first.", vbExclamation
            Exit Sub
        End If
    Next i
    
    newCount = 0
    ReDim newNodes(nodeCount - 2)
    
    For i = 0 To nodeCount - 1
        If i <> selectedIndex Then
            newNodes(newCount) = nodes(i)
            
            If newNodes(newCount).parentIndex > selectedIndex Then
                newNodes(newCount).parentIndex = newNodes(newCount).parentIndex - 1
            End If
            
            newCount = newCount + 1
        End If
    Next i
    
    For i = 0 To newCount - 1
        If newNodes(i).parentIndex > selectedIndex Then
            newNodes(i).parentIndex = newNodes(i).parentIndex - 1
        End If
    Next i
    
    nodes = newNodes
    nodeCount = newCount
    selectedIndex = -1
    txtNodeText.text = ""
    
    Call RefreshList
End Sub

Private Sub btnMoveUp_Click()
    Dim temp As PyramidNode
    Dim swapIdx As Long
    
    If selectedIndex <= 0 Then Exit Sub
    If selectedIndex = 1 Then Exit Sub
    
    swapIdx = FindPreviousSibling(selectedIndex)
    
    If swapIdx = -1 Then
        Exit Sub
    End If
    
    Call MoveBlockUp(selectedIndex, swapIdx)
    
    Dim swapBlockSize As Long
    swapBlockSize = GetBlockSize(swapIdx)
    selectedIndex = selectedIndex - swapBlockSize
    
    Call RefreshList
End Sub

Private Sub btnMoveDown_Click()
    Dim temp As PyramidNode
    Dim swapIdx As Long
    
    If selectedIndex < 0 Then Exit Sub
    If selectedIndex >= nodeCount - 1 Then Exit Sub
    If selectedIndex = 0 Then Exit Sub
    

    swapIdx = FindNextSibling(selectedIndex)
    
    If swapIdx = -1 Then

        Exit Sub
    End If
    

    Call MoveBlockDown(selectedIndex, swapIdx)
    

    Dim swapBlockSize As Long
    swapBlockSize = GetBlockSize(swapIdx)
    selectedIndex = selectedIndex + swapBlockSize
    
    Call RefreshList
End Sub

Private Function FindPreviousSibling(idx As Long) As Long

    Dim i As Long
    Dim targetDepth As Long
    Dim targetParent As Long
    Dim candidateIdx As Long
    
    targetDepth = nodes(idx).depth
    targetParent = nodes(idx).parentIndex
    
    i = idx - 1
    
    Do While i >= 0

        If nodes(i).depth > targetDepth Then
            i = i - 1

        ElseIf nodes(i).depth = targetDepth Then

            If nodes(i).parentIndex = targetParent Then
                FindPreviousSibling = i
                Exit Function
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    FindPreviousSibling = -1
End Function

Private Function FindNextSibling(idx As Long) As Long
    Dim i As Long
    Dim targetDepth As Long
    Dim targetParent As Long
    Dim blockSize As Long
    
    targetDepth = nodes(idx).depth
    targetParent = nodes(idx).parentIndex
    
    blockSize = GetBlockSize(idx)
    
    For i = idx + blockSize To nodeCount - 1
        If nodes(i).depth = targetDepth And nodes(i).parentIndex = targetParent Then
            FindNextSibling = i
            Exit Function
        End If
        
        If nodes(i).depth < targetDepth Then
            Exit For
        End If
    Next i
    
    FindNextSibling = -1
End Function

Private Function GetBlockSize(idx As Long) As Long
    Dim count As Long
    Dim i As Long
    Dim baseDepth As Long
    
    baseDepth = nodes(idx).depth
    count = 1
    
    For i = idx + 1 To nodeCount - 1
        If nodes(i).depth > baseDepth Then
            count = count + 1
        Else
            Exit For
        End If
    Next i
    
    GetBlockSize = count
End Function

Private Sub MoveBlockUp(sourceIdx As Long, targetIdx As Long)
    Dim sourceBlockSize As Long
    Dim targetBlockSize As Long
    Dim tempNodes() As PyramidNode
    Dim i As Long
    Dim writePos As Long
    
    sourceBlockSize = GetBlockSize(sourceIdx)
    targetBlockSize = GetBlockSize(targetIdx)
    
    ReDim tempNodes(nodeCount - 1)
    writePos = 0
    
    For i = 0 To targetIdx - 1
        tempNodes(writePos) = nodes(i)
        writePos = writePos + 1
    Next i
    
    For i = sourceIdx To sourceIdx + sourceBlockSize - 1
        tempNodes(writePos) = nodes(i)
        writePos = writePos + 1
    Next i
    
    For i = targetIdx To targetIdx + targetBlockSize - 1
        tempNodes(writePos) = nodes(i)
        writePos = writePos + 1
    Next i
    
    For i = targetIdx + targetBlockSize To sourceIdx - 1
        tempNodes(writePos) = nodes(i)
        writePos = writePos + 1
    Next i
    
    For i = sourceIdx + sourceBlockSize To nodeCount - 1
        tempNodes(writePos) = nodes(i)
        writePos = writePos + 1
    Next i
    
    nodes = tempNodes
    
    Call RecalculateParentIndices
End Sub

Private Sub MoveBlockDown(sourceIdx As Long, targetIdx As Long)

    Dim sourceBlockSize As Long
    Dim targetBlockSize As Long
    Dim tempNodes() As PyramidNode
    Dim i As Long
    Dim writePos As Long
    
    sourceBlockSize = GetBlockSize(sourceIdx)
    targetBlockSize = GetBlockSize(targetIdx)
    
    ReDim tempNodes(nodeCount - 1)
    writePos = 0
    
    For i = 0 To sourceIdx - 1
        tempNodes(writePos) = nodes(i)
        writePos = writePos + 1
    Next i
    
    For i = sourceIdx + sourceBlockSize To targetIdx - 1
        tempNodes(writePos) = nodes(i)
        writePos = writePos + 1
    Next i
    
    For i = targetIdx To targetIdx + targetBlockSize - 1
        tempNodes(writePos) = nodes(i)
        writePos = writePos + 1
    Next i
    
    For i = sourceIdx To sourceIdx + sourceBlockSize - 1
        tempNodes(writePos) = nodes(i)
        writePos = writePos + 1
    Next i
    
    For i = targetIdx + targetBlockSize To nodeCount - 1
        tempNodes(writePos) = nodes(i)
        writePos = writePos + 1
    Next i
    
    nodes = tempNodes
    
    Call RecalculateParentIndices
End Sub

Private Sub RecalculateParentIndices()
    Dim i As Long
    Dim j As Long
    Dim parentText As String
    Dim parentDepth As Long
    
    For i = 1 To nodeCount - 1
        parentDepth = nodes(i).depth - 1
        
        For j = i - 1 To 0 Step -1
            If nodes(j).depth = parentDepth Then
                nodes(i).parentIndex = j
                Exit For
            End If
        Next j
    Next i
End Sub

Private Sub btnGenerate_Click()
    Dim createNew As Boolean
    
    If Trim(txtSituation.text) = "" Or Trim(txtComplication.text) = "" Or Trim(txtQuestion.text) = "" Then
        MsgBox "Please fill in Situation, Complication, and Question.", vbExclamation
        Exit Sub
    End If
    
    If nodeCount = 0 Then
        MsgBox "Please create a pyramid structure.", vbExclamation
        Exit Sub
    End If
    
    Call SavePyramidToTags
    
    createNew = optNewPres.Value
    Call GenerateSlides(createNew)
    
    MsgBox "Slides generated successfully!", vbInformation
    
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub AddNode(text As String, depth As Long, parentIndex As Long)
    If nodeCount = 0 Then
        ReDim nodes(0)
    Else
        ReDim Preserve nodes(nodeCount)
    End If
    
    nodes(nodeCount).text = text
    nodes(nodeCount).depth = depth
    nodes(nodeCount).parentIndex = parentIndex
    
    nodeCount = nodeCount + 1
End Sub

Private Sub RefreshList()
    Dim i As Long
    Dim indent As String
    Dim displayText As String
    
    lstPyramid.Clear
    
    For i = 0 To nodeCount - 1
        indent = String(nodes(i).depth * 2, " ")
        If nodes(i).depth > 0 Then
            indent = indent & "- "
        End If
        
        displayText = indent & nodes(i).text
        lstPyramid.AddItem displayText
    Next i
    
    If selectedIndex >= 0 And selectedIndex < nodeCount Then
        lstPyramid.ListIndex = selectedIndex
    End If
End Sub

Private Sub LoadPyramidFromTags()
    On Error Resume Next
    
    Dim structureData As String
    Dim items() As String
    Dim parts() As String
    Dim i As Long
    Dim sld As Slide
    Dim slideIdx As Long
    
    txtSituation.text = ActivePresentation.Tags("InstrumentaPyramidSCQ_Situation")
    txtComplication.text = ActivePresentation.Tags("InstrumentaPyramidSCQ_Complication")
    txtQuestion.text = ActivePresentation.Tags("InstrumentaPyramidSCQ_Question")
    
    structureData = ActivePresentation.Tags("InstrumentaPyramidSCQ_Structure")
    
    If structureData <> "" Then
        items = Split(structureData, ";")
        nodeCount = 0
        ReDim nodes(UBound(items))
        
        For i = 0 To UBound(items)
            parts = Split(items(i), "|")
            If UBound(parts) >= 2 Then
                nodes(nodeCount).text = parts(0)
                nodes(nodeCount).depth = CLng(parts(1))
                nodes(nodeCount).parentIndex = CLng(parts(2))
                nodeCount = nodeCount + 1
            End If
        Next i
        
        For i = 1 To nodeCount - 1
            slideIdx = i + 2
            Set sld = FindPyramidSlide(ActivePresentation, slideIdx)
            
            If Not sld Is Nothing Then
                On Error Resume Next
                If sld.Shapes.HasTitle Then
                    nodes(i).text = sld.Shapes.Title.TextFrame.textRange.text
                End If
                On Error GoTo 0
            End If
        Next i
    End If
    
    On Error GoTo 0
End Sub

Private Sub SavePyramidToTags()
    Dim structureData As String
    Dim i As Long
    
    ActivePresentation.Tags.Add "InstrumentaPyramidSCQ_Situation", Trim(txtSituation.text)
    ActivePresentation.Tags.Add "InstrumentaPyramidSCQ_Complication", Trim(txtComplication.text)
    ActivePresentation.Tags.Add "InstrumentaPyramidSCQ_Question", Trim(txtQuestion.text)
    
    structureData = ""
    For i = 0 To nodeCount - 1
        If i > 0 Then structureData = structureData & ";"
        structureData = structureData & nodes(i).text & "|" & nodes(i).depth & "|" & nodes(i).parentIndex
    Next i
    
    ActivePresentation.Tags.Add "InstrumentaPyramidSCQ_Structure", structureData
End Sub

Private Function FindPyramidSlide(pres As Presentation, slideIndex As Long) As Slide
    Dim sld As Slide
    
    On Error Resume Next
    For Each sld In pres.Slides
        If sld.Tags("InstrumentaPyramidSlideIndex") = CStr(slideIndex) Then
            Set FindPyramidSlide = sld
            Exit Function
        End If
    Next sld
    On Error GoTo 0
    
    Set FindPyramidSlide = Nothing
End Function

Private Sub GenerateSlides(createNew As Boolean)
    Dim pres As Presentation
    Dim sld As Slide
    Dim shp As shape
    Dim i As Long
    Dim leftPos As Single
    Dim colWidth As Single
    Dim slideIdx As Long
    Dim targetPosition As Long
    
    If createNew Then
        Set pres = Presentations.Add(msoTrue)
    Else
        Set pres = ActivePresentation
    End If
    
    slideIdx = 0
    targetPosition = pres.Slides.count + 1 '
    
    slideIdx = slideIdx + 1
    Set sld = FindPyramidSlide(pres, slideIdx)
    
    If sld Is Nothing Then
        Set sld = pres.Slides.Add(targetPosition, ppLayoutBlank)
        sld.Tags.Add "InstrumentaPyramidSlideIndex", CStr(slideIdx)
        sld.layout = ppLayoutBlank
        
        Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 30, pres.PageSetup.slideWidth - 100, 60)
        shp.Tags.Add "InstrumentaPyramidElement", "Title"
        With shp.TextFrame2.textRange
            .text = "Situation - Complication - Question"
            .Font.Size = 32
            .Font.Bold = msoTrue
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
        
        colWidth = (pres.PageSetup.slideWidth - 150) / 3
        
        Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 120, colWidth, 300)
        shp.Tags.Add "InstrumentaPyramidElement", "Situation"
        
        leftPos = 50 + colWidth + 25
        Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, 120, colWidth, 300)
        shp.Tags.Add "InstrumentaPyramidElement", "Complication"
        
        leftPos = leftPos + colWidth + 25
        Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, 120, colWidth, 300)
        shp.Tags.Add "InstrumentaPyramidElement", "Question"
        
        targetPosition = targetPosition + 1
    Else
        Call MoveSlideToPosition(sld, slideIdx, pres)
    End If
    
    For Each shp In sld.Shapes
        If shp.Tags("InstrumentaPyramidElement") = "Situation" Then
            With shp
                .TextFrame2.WordWrap = msoTrue
                .TextFrame2.textRange.text = "Situation" & vbCrLf & vbCrLf & txtSituation.text
                .TextFrame2.textRange.Font.Size = 14
                .TextFrame2.textRange.Paragraphs(1).Font.Bold = msoTrue
                .TextFrame2.textRange.Paragraphs(1).Font.Size = 18
            End With
        ElseIf shp.Tags("InstrumentaPyramidElement") = "Complication" Then
            With shp
                .TextFrame2.WordWrap = msoTrue
                .TextFrame2.textRange.text = "Complication" & vbCrLf & vbCrLf & txtComplication.text
                .TextFrame2.textRange.Font.Size = 14
                .TextFrame2.textRange.Paragraphs(1).Font.Bold = msoTrue
                .TextFrame2.textRange.Paragraphs(1).Font.Size = 18
            End With
        ElseIf shp.Tags("InstrumentaPyramidElement") = "Question" Then
            With shp
                .TextFrame2.WordWrap = msoTrue
                .TextFrame2.textRange.text = "Question" & vbCrLf & vbCrLf & txtQuestion.text
                .TextFrame2.textRange.Font.Size = 14
                .TextFrame2.textRange.Paragraphs(1).Font.Bold = msoTrue
                .TextFrame2.textRange.Paragraphs(1).Font.Size = 18
            End With
        End If
    Next shp
    
    slideIdx = slideIdx + 1
    Set sld = FindPyramidSlide(pres, slideIdx)
    
    If sld Is Nothing Then
        Set sld = pres.Slides.Add(targetPosition, ppLayoutText)
        sld.Tags.Add "InstrumentaPyramidSlideIndex", CStr(slideIdx)
        targetPosition = targetPosition + 1
    Else
        Call MoveSlideToPosition(sld, slideIdx, pres)
    End If
    
    sld.Shapes.Title.TextFrame.textRange.text = "Management Summary"
    
    Dim summaryText As String
    Dim tr As TextRange2
    Dim para As Long
    
    summaryText = nodes(0).text & vbCrLf & vbCrLf
    
    For i = 1 To nodeCount - 1
        If nodes(i).depth = 1 Then
            summaryText = summaryText & nodes(i).text & vbCrLf
            
            Call AddChildrenToSummary(summaryText, i, 2)
        End If
    Next i
    
    Set tr = sld.Shapes.Placeholders(2).TextFrame2.textRange
    tr.text = summaryText
    
    para = 1
    For i = 3 To tr.Paragraphs.count
        If para <= nodeCount Then
            Dim nodeIdx As Long
            nodeIdx = GetNodeIndexForParagraph(i - 2)
            
            If nodeIdx > 0 And nodeIdx < nodeCount Then
                With tr.Paragraphs(i).ParagraphFormat
                    .Bullet.visible = msoTrue
                    .Bullet.Type = msoBulletUnnumbered
                    
                    If nodes(nodeIdx).depth = 1 Then
                        .IndentLevel = 1
                        .LeftIndent = 0
                    ElseIf nodes(nodeIdx).depth = 2 Then
                        .IndentLevel = 2
                        .LeftIndent = 36
                    End If
                End With
            End If
        End If
        para = para + 1
    Next i
    
    For i = 1 To nodeCount - 1
        slideIdx = slideIdx + 1
        Set sld = FindPyramidSlide(pres, slideIdx)
        
        If sld Is Nothing Then
            Set sld = pres.Slides.Add(targetPosition, ppLayoutText)
            sld.Tags.Add "InstrumentaPyramidSlideIndex", CStr(slideIdx)
            targetPosition = targetPosition + 1
        Else
            Call MoveSlideToPosition(sld, slideIdx, pres)
        End If
        
        sld.Shapes.Title.TextFrame.textRange.text = nodes(i).text
    Next i
    
    Call DeleteOrphanPyramidSlides(pres, slideIdx)
    
End Sub

Private Sub MoveSlideToPosition(sld As Slide, expectedIndex As Long, pres As Presentation)
    
    Dim currentPos As Long
    Dim targetPos As Long
    Dim s As Slide
    Dim pyramidSlidesBefore As Long
    
    currentPos = sld.slideIndex
    
    pyramidSlidesBefore = 0
    For Each s In pres.Slides
        If s.Tags("InstrumentaPyramidSlideIndex") <> "" Then
            If CLng(s.Tags("InstrumentaPyramidSlideIndex")) < expectedIndex Then
                pyramidSlidesBefore = pyramidSlidesBefore + 1
            End If
        End If
    Next s
    
    Dim firstPyramidPos As Long
    firstPyramidPos = 999999
    
    For Each s In pres.Slides
        If s.Tags("InstrumentaPyramidSlideIndex") = "1" Then
            firstPyramidPos = s.slideIndex
            Exit For
        End If
    Next s
    
    If firstPyramidPos = 999999 Then
        targetPos = pres.Slides.count
    Else
        targetPos = firstPyramidPos + pyramidSlidesBefore
    End If
    
    If currentPos <> targetPos Then
        sld.MoveTo targetPos
    End If
End Sub

Private Sub DeleteOrphanPyramidSlides(pres As Presentation, maxValidIndex As Long)
    Dim sld As Slide
    Dim i As Long
    
    For i = pres.Slides.count To 1 Step -1
        Set sld = pres.Slides(i)
        If sld.Tags("InstrumentaPyramidSlideIndex") <> "" Then
            If CLng(sld.Tags("InstrumentaPyramidSlideIndex")) > maxValidIndex Then
                sld.Delete
            End If
        End If
    Next i
End Sub

Private Function GetNodeIndexForParagraph(paraNum As Long) As Long
    Dim count As Long
    Dim i As Long
    
    count = 0
    For i = 1 To nodeCount - 1
        If nodes(i).depth <= 2 Then
            count = count + 1
            If count = paraNum Then
                GetNodeIndexForParagraph = i
                Exit Function
            End If
        End If
    Next i
    
    GetNodeIndexForParagraph = -1
End Function

Private Sub AddChildrenToSummary(ByRef summaryText As String, parentIndex As Long, maxDepth As Long)
    Dim i As Long
    
    For i = 0 To nodeCount - 1
        If nodes(i).parentIndex = parentIndex And nodes(i).depth <= maxDepth Then
            summaryText = summaryText & nodes(i).text & vbCrLf
        End If
    Next i
End Sub

Private Sub RemoveAllInstrumentaPyramidTags()

    Dim pres As Presentation
    Dim sld As Slide
    Dim shp As shape
    
    Set pres = ActivePresentation

    On Error Resume Next
    pres.Tags.Delete "InstrumentaPyramidSCQ_Situation"
    pres.Tags.Delete "InstrumentaPyramidSCQ_Complication"
    pres.Tags.Delete "InstrumentaPyramidSCQ_Question"
    pres.Tags.Delete "InstrumentaPyramidSCQ_Structure"
    On Error GoTo 0

    For Each sld In pres.Slides
        On Error Resume Next
        sld.Tags.Delete "InstrumentaPyramidSlideIndex"
        On Error GoTo 0
        
        For Each shp In sld.Shapes
            On Error Resume Next
            shp.Tags.Delete "InstrumentaPyramidElement"
            On Error GoTo 0
        Next shp
    Next sld

    MsgBox "All Instrumenta Pyramid tags have been removed.", _
           vbInformation, "Cleanup Complete"

End Sub

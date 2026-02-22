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

#If Mac Then
#Else
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal cb As LongPtr)
#End If

Const GMEM_MOVEABLE As Long = &H2
Const CF_UNICODETEXT As Long = 13
#End If

Private Type PyramidNode
    text As String
    depth As Long
    parentIndex As Long
End Type

Private nodes() As PyramidNode
Private nodeCount As Long
Private selectedIndex As Long

Private Sub btnClear_Click()
    Dim answer As VbMsgBoxResult
    
    answer = MsgBox("This will delete all pyramid data and cannot be undone." & vbCrLf & vbCrLf & _
                    "Are you sure you want to continue?", _
                    vbYesNo + vbExclamation, "Confirm Clear")
    
    If answer = vbNo Then Exit Sub
    
    RemoveAllInstrumentaPyramidTags
    Unload Me
End Sub

Private Function FindListIndexForNode(nodeIdx As Long) As Long
    Dim i As Long
    For i = 0 To lstPyramid.ListCount - 1
        If CLng(lstPyramid.List(i, 1)) = nodeIdx Then
            FindListIndexForNode = i
            Exit Function
        End If
    Next i
    FindListIndexForNode = -1
End Function




Public Sub CopyToClipboard(text As String)

    #If Mac Then
   
    MacScript "set the clipboard to """ & Replace(text, """", "\""") & """"
    
    #Else

    Dim hMem As LongPtr
    Dim pMem As LongPtr
    Dim bytes As Long
    
    bytes = (Len(text) * 2) + 2

    OpenClipboard 0
    EmptyClipboard

    hMem = GlobalAlloc(GMEM_MOVEABLE, bytes)
    pMem = GlobalLock(hMem)

    CopyMemory pMem, StrPtr(text), bytes

    GlobalUnlock hMem
    SetClipboardData CF_UNICODETEXT, hMem
    CloseClipboard

    #End If
    
End Sub


Private Function PromptJsonInput(title As String) As String
    With JsonInputForm
        .Caption = title
        .txtJson.text = ""
        .Cancelled = False
        .Show
        
        If .Cancelled Then
            PromptJsonInput = ""
        Else
            PromptJsonInput = .result
        End If
    End With
End Function

Private Sub btnImportAI_Click()
    Dim jsonText As String
    
    jsonText = PromptJsonInput("Paste AI JSON Result")
    If Trim(jsonText) = "" Then Exit Sub
    
    On Error GoTo ParseError
    Call ParsePyramidJson(jsonText)
    Call RefreshList
    
    MsgBox "AI JSON imported successfully!", vbInformation
    Exit Sub

ParseError:
    MsgBox "The JSON could not be parsed. Please check the format.", vbCritical
End Sub

Private Sub btnImproveStorylinePrompt_Click()

    Dim answer As VbMsgBoxResult
    Dim prompt As String
    Dim json As String
    Dim i As Long

    answer = MsgBox( _
        "This action will include the full current storyline (SCQ + all nodes) in the AI prompt." & vbCrLf & vbCrLf & _
        "If your storyline contains confidential, personal, or sensitive information, " & _
        "be aware that sending it to an external AI service may pose a security or privacy risk." & vbCrLf & vbCrLf & _
        "Do you want to continue?", _
        vbYesNo + vbExclamation, _
        "Security Warning")

    If answer = vbNo Then Exit Sub

    json = "{""situation"": """ & EscapeJson(txtSituation.text) & """," & vbCrLf
    json = json & " ""complication"": """ & EscapeJson(txtComplication.text) & """," & vbCrLf
    json = json & " ""question"": """ & EscapeJson(txtQuestion.text) & """," & vbCrLf
    json = json & " ""nodes"": [" & vbCrLf

    For i = 0 To nodeCount - 1
        json = json & "    {""text"": """ & EscapeJson(nodes(i).text) & """, "
        json = json & """depth"": " & nodes(i).depth & ", "
        json = json & """parentIndex"": " & nodes(i).parentIndex & "}"
        If i < nodeCount - 1 Then json = json & ","
        json = json & vbCrLf
    Next i

    json = json & "  ]" & vbCrLf & "}" & vbCrLf

    prompt = ""
    prompt = prompt & "I want to improve an existing Pyramid Principle storyline." & vbCrLf & vbCrLf
    prompt = prompt & "Below is the current storyline in JSON format. Please analyze it carefully:" & vbCrLf & vbCrLf
    prompt = prompt & json & vbCrLf

    prompt = prompt & "Your task is to help me refine, strengthen, and improve this storyline." & vbCrLf & vbCrLf

    prompt = prompt & "Please follow these instructions exactly:" & vbCrLf & vbCrLf

    prompt = prompt & "1. First, ask me what I want to improve" & vbCrLf
    prompt = prompt & "   - clarity" & vbCrLf
    prompt = prompt & "   - logic" & vbCrLf
    prompt = prompt & "   - MECE structure" & vbCrLf
    prompt = prompt & "   - slide titles" & vbCrLf
    prompt = prompt & "   - depth/structure" & vbCrLf
    prompt = prompt & "   - or all of the above" & vbCrLf & vbCrLf

    prompt = prompt & "2. After I specify what I want improved, propose a revised storyline:" & vbCrLf
    prompt = prompt & "   - Rewrite the Situation, Complication, and Question if needed" & vbCrLf
    prompt = prompt & "   - Improve the Answer" & vbCrLf
    prompt = prompt & "   - Rewrite key lines as full-sentence slide titles" & vbCrLf
    prompt = prompt & "   - Ensure all supporting points are full-sentence and logically grouped" & vbCrLf
    prompt = prompt & "   - Ensure the structure is MECE and follows the Pyramid Principle" & vbCrLf & vbCrLf

    prompt = prompt & "3. Present the improved storyline as a hierarchical bullet list using this exact formatting:" & vbCrLf & vbCrLf
    prompt = prompt & "- **Answer:** Full-sentence answer" & vbCrLf
    prompt = prompt & "  - Full-sentence key line" & vbCrLf
    prompt = prompt & "    - Full-sentence supporting point" & vbCrLf
    prompt = prompt & "    - Full-sentence supporting point" & vbCrLf
    prompt = prompt & "  - Another full-sentence key line" & vbCrLf
    prompt = prompt & "    - Supporting point" & vbCrLf
    prompt = prompt & "    - Supporting point" & vbCrLf & vbCrLf

    prompt = prompt & "(Use hyphens for bullets and two spaces for indentation per level.)" & vbCrLf & vbCrLf

    prompt = prompt & "4. Ask me whether I want further changes or if I approve the improved storyline." & vbCrLf & vbCrLf

    prompt = prompt & "5. ONLY AFTER I approve the improved storyline, generate the final JSON output in this exact format:" & vbCrLf & vbCrLf

    prompt = prompt & "{" & vbCrLf
    prompt = prompt & "  ""situation"": ""…""," & vbCrLf
    prompt = prompt & "  ""complication"": ""…""," & vbCrLf
    prompt = prompt & "  ""question"": ""…""," & vbCrLf
    prompt = prompt & "  ""nodes"": [" & vbCrLf
    prompt = prompt & "    {""text"": ""Answer sentence"", ""depth"": 0, ""parentIndex"": -1}," & vbCrLf
    prompt = prompt & "    {""text"": ""Full-sentence key line"", ""depth"": 1, ""parentIndex"": 0}," & vbCrLf
    prompt = prompt & "    {""text"": ""Full-sentence supporting point"", ""depth"": 2, ""parentIndex"": 1}" & vbCrLf
    prompt = prompt & "  ]" & vbCrLf
    prompt = prompt & "}" & vbCrLf & vbCrLf

    prompt = prompt & "6. JSON rules you MUST follow:" & vbCrLf
    prompt = prompt & "   - The first node must always be the Answer node with depth 0 and parentIndex -1." & vbCrLf
    prompt = prompt & "   - Every node must include:" & vbCrLf
    prompt = prompt & "       ""text"": string (full-sentence slide title, no ending period)" & vbCrLf
    prompt = prompt & "       ""depth"": integer (0 = top, increasing for children)" & vbCrLf
    prompt = prompt & "       ""parentIndex"": integer (index of parent node in the array)" & vbCrLf
    prompt = prompt & "   - parentIndex must always point to a node that appears earlier in the list." & vbCrLf
    prompt = prompt & "   - The structure must form a valid tree." & vbCrLf
    prompt = prompt & "   - No trailing commas." & vbCrLf
    prompt = prompt & "   - No commentary outside the JSON." & vbCrLf & vbCrLf

    prompt = prompt & "After you generate the JSON, stop. Do not add explanations or notes." & vbCrLf

    CopyToClipboard prompt

    MsgBox "The improvement prompt has been copied to your clipboard.", vbInformation

End Sub




Private Sub btnStorylinePrompt_Click()


    Dim prompt As String
    Dim shp As shape
    Dim sld As Slide

    prompt = ""
    prompt = prompt & "I want to generate a Pyramid Principle storyline in JSON format so I can import it into my tool." & vbCrLf & vbCrLf
    prompt = prompt & "Please follow these instructions exactly:" & vbCrLf & vbCrLf

    prompt = prompt & "1. First, ask me for the topic or business situation I want the storyline to be about." & vbCrLf & vbCrLf

    prompt = prompt & "2. After I provide the topic, create:" & vbCrLf
    prompt = prompt & "   - A clear Situation (S)" & vbCrLf
    prompt = prompt & "   - A clear Complication (C)" & vbCrLf
    prompt = prompt & "   - A clear Question (Q)" & vbCrLf
    prompt = prompt & "   - A draft storyline written as a hierarchical bullet list." & vbCrLf & vbCrLf

    prompt = prompt & "3. Every bullet in the draft storyline must:" & vbCrLf
    prompt = prompt & "   - Be a full-sentence slide title" & vbCrLf
    prompt = prompt & "   - Clearly state the main conclusion of that slide (consulting-style key message)" & vbCrLf
    prompt = prompt & "   - NOT end with a period or any other punctuation mark" & vbCrLf
    prompt = prompt & "   - NOT be a fragment or label" & vbCrLf & vbCrLf

    prompt = prompt & "4. Present the draft storyline using this exact bullet formatting:" & vbCrLf & vbCrLf
    prompt = prompt & "- **Answer:** Full-sentence answer" & vbCrLf
    prompt = prompt & "  - Full-sentence key line" & vbCrLf
    prompt = prompt & "    - Full-sentence supporting point" & vbCrLf
    prompt = prompt & "    - Full-sentence supporting point" & vbCrLf
    prompt = prompt & "  - Another full-sentence key line" & vbCrLf
    prompt = prompt & "    - Supporting point" & vbCrLf
    prompt = prompt & "    - Supporting point" & vbCrLf & vbCrLf

    prompt = prompt & "(Use hyphens for bullets and two spaces for indentation per level.)" & vbCrLf & vbCrLf

    prompt = prompt & "5. After presenting the draft storyline, ask me whether I want to:" & vbCrLf
    prompt = prompt & "   - Approve it as-is, or" & vbCrLf
    prompt = prompt & "   - Request changes or improvements" & vbCrLf & vbCrLf

    prompt = prompt & "6. Only AFTER I approve the storyline, generate the final output in the following JSON format:" & vbCrLf & vbCrLf

    prompt = prompt & "{" & vbCrLf
    prompt = prompt & "  ""situation"": ""…""," & vbCrLf
    prompt = prompt & "  ""complication"": ""…""," & vbCrLf
    prompt = prompt & "  ""question"": ""…""," & vbCrLf
    prompt = prompt & "  ""nodes"": [" & vbCrLf
    prompt = prompt & "    {""text"": ""Answer sentence"", ""depth"": 0, ""parentIndex"": -1}," & vbCrLf
    prompt = prompt & "    {""text"": ""Full-sentence key line"", ""depth"": 1, ""parentIndex"": 0}," & vbCrLf
    prompt = prompt & "    {""text"": ""Full-sentence supporting point"", ""depth"": 2, ""parentIndex"": 1}" & vbCrLf
    prompt = prompt & "  ]" & vbCrLf
    prompt = prompt & "}" & vbCrLf & vbCrLf

    prompt = prompt & "7. JSON rules you MUST follow:" & vbCrLf
    prompt = prompt & "   - The first node must always be the Answer node with depth 0 and parentIndex -1." & vbCrLf
    prompt = prompt & "   - Every node must include:" & vbCrLf
    prompt = prompt & "       ""text"": string (full-sentence slide title, no ending period)" & vbCrLf
    prompt = prompt & "       ""depth"": integer (0 = top, increasing for children)" & vbCrLf
    prompt = prompt & "       ""parentIndex"": integer (index of parent node in the array)" & vbCrLf
    prompt = prompt & "   - parentIndex must always point to a node that appears earlier in the list." & vbCrLf
    prompt = prompt & "   - The structure must form a valid tree." & vbCrLf
    prompt = prompt & "   - No trailing commas." & vbCrLf
    prompt = prompt & "   - No commentary outside the JSON." & vbCrLf & vbCrLf

    prompt = prompt & "8. Unless I specify otherwise:" & vbCrLf
    prompt = prompt & "   - You may generate any number of nodes." & vbCrLf
    prompt = prompt & "   - Ensure the storyline is coherent, MECE, and follows the Pyramid Principle." & vbCrLf & vbCrLf

    prompt = prompt & "After you generate the JSON, stop. Do not add explanations or notes." & vbCrLf

   CopyToClipboard prompt

   MsgBox "The storyline prompt has been copied to your clipboard.", vbInformation


End Sub


Private Sub UserForm_Initialize()
    selectedIndex = -1
    nodeCount = 0
        
    optCurrentPres.value = True
    
    With lstPyramid
    .ColumnCount = 2
    .ColumnWidths = "200 pt;0 pt"
    .BoundColumn = 1
    End With

    
    Call LoadPyramidFromTags
    
    If nodeCount = 0 Then
        If ActivePresentation.Slides.count > 0 Then
            Dim answer As VbMsgBoxResult
            answer = MsgBox("Would you like to import the current slide titles as your pyramid structure?" & vbCrLf & vbCrLf & _
                           "This will create a pyramid from existing slide titles (under 'Answer' node).", _
                           vbYesNo + vbQuestion, "Import Existing Storyline")
            
            If answer = vbYes Then
                Call ImportExistingSlides
            Else
                Call AddNode("Answer", 0, -1)
                Call AddNode("Key Line 1", 1, 0)
                Call AddNode("Key Line 2", 1, 0)
            End If
        Else
            Call AddNode("Answer", 0, -1)
            Call AddNode("Key Line 1", 1, 0)
            Call AddNode("Key Line 2", 1, 0)
        End If
    End If
    
    Call RefreshList
End Sub

Private Sub ImportExistingSlides()
    Dim sld As Slide
    Dim slideTitle As String
    Dim pyramidSlideIdx As Long
    Dim contentSlides As Collection
    Dim i As Long
    
    Call AddNode("Answer", 0, -1)
    
    Set contentSlides = New Collection
    
    For Each sld In ActivePresentation.Slides
        If sld.layout <> ppLayoutTitle And sld.layout <> ppLayoutTitleOnly Then

            On Error Resume Next
            slideTitle = ""
            If sld.shapes.HasTitle Then
                slideTitle = Trim(sld.shapes.title.TextFrame.textRange.text)
            End If
            On Error GoTo 0
            

            If slideTitle <> "" Then
               
                Call AddNode(slideTitle, 1, 0)
                
  
                contentSlides.Add sld
            End If
        End If
    Next sld
    

    pyramidSlideIdx = 3
    
    For i = 1 To contentSlides.count
        Set sld = contentSlides(i)
        sld.Tags.Add "InstrumentaPyramidSlideIndex", CStr(pyramidSlideIdx)
        pyramidSlideIdx = pyramidSlideIdx + 1
    Next i
    

    If nodeCount = 1 Then
        Call AddNode("Key Line 1", 1, 0)
        Call AddNode("Key Line 2", 1, 0)
    End If
    
    MsgBox nodeCount - 1 & " slide title(s) imported as pyramid structure." & vbCrLf & _
           "(Title slides were skipped, existing slides will be reused)", vbInformation
End Sub

Private Sub lstPyramid_Click()
    If lstPyramid.ListIndex >= 0 Then
        selectedIndex = CLng(lstPyramid.List(lstPyramid.ListIndex, 1))
        txtNodeText.text = nodes(selectedIndex).text
    End If
End Sub

Private Function GetInsertIndexForChild(parentIdx As Long) As Long
    Dim i As Long
    Dim baseDepth As Long
    
    baseDepth = nodes(parentIdx).depth
    
    GetInsertIndexForChild = parentIdx + 1
    
    For i = parentIdx + 1 To nodeCount - 1
        If nodes(i).depth <= baseDepth Then
            Exit Function
        End If
        GetInsertIndexForChild = i + 1
    Next i
End Function

Private Sub InsertNodeAt(insertIdx As Long, text As String, depth As Long, parentIndex As Long)
    Dim i As Long
    
    If nodeCount = 0 Then
        ReDim nodes(0)
    Else
        ReDim Preserve nodes(nodeCount)
        For i = nodeCount - 1 To insertIdx Step -1
            nodes(i + 1) = nodes(i)
        Next i
    End If
    
    nodes(insertIdx).text = text
    nodes(insertIdx).depth = depth
    nodes(insertIdx).parentIndex = parentIndex
    
    nodeCount = nodeCount + 1
End Sub



Private Sub btnAddChild_Click()
    Dim newText As String
    Dim insertIdx As Long
    
    If selectedIndex < 0 Then
        MsgBox "Please select a parent node first.", vbExclamation
        Exit Sub
    End If
    
    newText = Trim(txtNodeText.text)
    If newText = "" Then
        MsgBox "Please enter text for the new node.", vbExclamation
        Exit Sub
    End If
    
    insertIdx = GetInsertIndexForChild(selectedIndex)
    Call InsertNodeAt(insertIdx, newText, nodes(selectedIndex).depth + 1, selectedIndex)
    
    selectedIndex = insertIdx
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
    Dim previewMsg As String
    Dim answer As VbMsgBoxResult
    Dim newCount As Long
    Dim updateCount As Long
    Dim totalSlides As Long
    
    If Trim(txtSituation.text) = "" Or Trim(txtComplication.text) = "" Or Trim(txtQuestion.text) = "" Then
        MsgBox "Please fill in Situation, Complication, and Question.", vbExclamation
        Exit Sub
    End If
    
    If nodeCount = 0 Then
        MsgBox "Please create a pyramid structure.", vbExclamation
        Exit Sub
    End If
    
    createNew = optNewPres.value
    
    totalSlides = 2 + (nodeCount - 1)
    
    If createNew Then
        previewMsg = "Will create " & totalSlides & " new slides in a new presentation." & vbCrLf & vbCrLf & _
                     "Continue?"
    Else
        Call CountCreateVsUpdate(newCount, updateCount)
        
        If newCount > 0 And updateCount > 0 Then
            previewMsg = "Will create " & newCount & " new slide(s) and update " & updateCount & " existing slide(s)." & vbCrLf & vbCrLf & _
                         "Continue?"
        ElseIf newCount > 0 Then
            previewMsg = "Will create " & newCount & " new slide(s)." & vbCrLf & vbCrLf & _
                         "Continue?"
        Else
            previewMsg = "Will update " & updateCount & " existing slide(s)." & vbCrLf & vbCrLf & _
                         "Continue?"
        End If
    End If
    
    answer = MsgBox(previewMsg, vbYesNo + vbQuestion, "Generate Slides")
    If answer = vbNo Then Exit Sub
    
    Call SavePyramidToTags
    Call GenerateSlides(createNew)
    
    MsgBox "Slides generated successfully!", vbInformation
    
    Unload Me
End Sub

Private Sub CountCreateVsUpdate(ByRef newCount As Long, ByRef updateCount As Long)
    Dim i As Long
    Dim slideIdx As Long
    Dim sld As Slide
    
    newCount = 0
    updateCount = 0
    
    slideIdx = 1
    Set sld = FindPyramidSlide(ActivePresentation, slideIdx)
    If sld Is Nothing Then
        newCount = newCount + 1
    Else
        updateCount = updateCount + 1
    End If
    
    slideIdx = 2
    Set sld = FindPyramidSlide(ActivePresentation, slideIdx)
    If sld Is Nothing Then
        newCount = newCount + 1
    Else
        updateCount = updateCount + 1
    End If
    
    For i = 1 To nodeCount - 1
        slideIdx = i + 2
        Set sld = FindPyramidSlide(ActivePresentation, slideIdx)
        If sld Is Nothing Then
            newCount = newCount + 1
        Else
            updateCount = updateCount + 1
        End If
    Next i
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnExportPyramid_Click()
    Dim exportPath As String
    Dim fileNum As Integer
    Dim jsonContent As String
    Dim i As Long
    
    #If Mac Then
        exportPath = MacSaveAsDialog("PyramidStructure.json")
        If exportPath = "" Or exportPath = "False" Then Exit Sub
    #Else
        Dim fd As Object
        Set fd = Application.FileDialog(msoFileDialogSaveAs)
        With fd
            .title = "Export Pyramid Structure"
            .InitialFileName = "PyramidStructure.json"
            .FilterIndex = 1
            
            If .Show <> -1 Then Exit Sub
            exportPath = .SelectedItems(1)
        End With
    #End If
     
    
    If LCase(right(exportPath, 5)) <> ".json" Then
        exportPath = exportPath & ".json"
    End If
    
    jsonContent = "{" & vbCrLf
    jsonContent = jsonContent & "  ""situation"": """ & EscapeJson(txtSituation.text) & """," & vbCrLf
    jsonContent = jsonContent & "  ""complication"": """ & EscapeJson(txtComplication.text) & """," & vbCrLf
    jsonContent = jsonContent & "  ""question"": """ & EscapeJson(txtQuestion.text) & """," & vbCrLf
    jsonContent = jsonContent & "  ""nodes"": [" & vbCrLf
    
    For i = 0 To nodeCount - 1
        jsonContent = jsonContent & "    {""text"": """ & EscapeJson(nodes(i).text) & """, "
        jsonContent = jsonContent & """depth"": " & nodes(i).depth & ", "
        jsonContent = jsonContent & """parentIndex"": " & nodes(i).parentIndex & "}"
        If i < nodeCount - 1 Then jsonContent = jsonContent & ","
        jsonContent = jsonContent & vbCrLf
    Next i
    
    jsonContent = jsonContent & "  ]" & vbCrLf
    jsonContent = jsonContent & "}" & vbCrLf
    
    fileNum = FreeFile
    Open exportPath For Output As #fileNum
    Print #fileNum, jsonContent
    Close #fileNum
    
    MsgBox "Pyramid structure exported to:" & vbCrLf & exportPath, vbInformation
End Sub

Private Sub btnImportPyramid_Click()
    Dim importPath As String
    Dim fileNum As Integer
    Dim jsonContent As String
    Dim fileLine As String
    Dim answer As VbMsgBoxResult
    
    If nodeCount > 0 Then
        answer = MsgBox("Importing will replace the current pyramid structure." & vbCrLf & vbCrLf & _
                        "Continue?", vbYesNo + vbQuestion, "Confirm Import")
        If answer = vbNo Then Exit Sub
    End If
    
    #If Mac Then
        importPath = MacFileDialog("/")
        If importPath = "" Or importPath = "False" Then Exit Sub
    #Else
        Dim fd As Object
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .title = "Import Pyramid Structure"
            .Filters.Clear
            .Filters.Add "JSON Files", "*.json"
            
            If .Show <> -1 Then Exit Sub
            importPath = .SelectedItems(1)
        End With
    #End If
    
    fileNum = FreeFile
    Open importPath For Input As #fileNum
    jsonContent = ""
    Do Until EOF(fileNum)
        Line Input #fileNum, fileLine
        jsonContent = jsonContent & fileLine & vbCrLf
    Loop
    Close #fileNum
    
    Call ParsePyramidJson(jsonContent)
    
    Call RefreshList
    
    MsgBox "Pyramid structure imported successfully!", vbInformation
End Sub

Private Function EscapeJson(text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJson = result
End Function

Private Sub ParsePyramidJson(jsonContent As String)
    Dim cleaned As String
    Dim pos As Long
    Dim arrNodes As Variant
    Dim node As Variant
    
    cleaned = Replace(jsonContent, vbCrLf, vbLf)
    cleaned = Replace(cleaned, vbCr, vbLf)
    
    txtSituation.text = ExtractJsonValueByKey(cleaned, "situation")
    txtComplication.text = ExtractJsonValueByKey(cleaned, "complication")
    txtQuestion.text = ExtractJsonValueByKey(cleaned, "question")
    
    Dim nodesJson As String
    nodesJson = ExtractJsonArray(cleaned, "nodes")
    
    arrNodes = SplitNodes(nodesJson)
    
    nodeCount = 0
    ReDim nodes(0)
    
    For Each node In arrNodes
        Dim t As String, d As Long, p As Long
        
        t = ExtractJsonValueByKey(CStr(node), "text")
        d = CLng(ExtractJsonValueByKey(CStr(node), "depth"))
        p = CLng(ExtractJsonValueByKey(CStr(node), "parentIndex"))
        
        AddNode t, d, p
    Next node
End Sub

Private Function ExtractJsonValueByKey(ByVal json As String, key As String) As String
    Dim pattern As String
    Dim startPos As Long, endPos As Long
    
    pattern = """" & key & """:"
    
    startPos = InStr(1, json, pattern)
    If startPos = 0 Then Exit Function
    
    startPos = startPos + Len(pattern)
    
    Do While Mid$(json, startPos, 1) = " "
        startPos = startPos + 1
    Loop
    
    If Mid$(json, startPos, 1) = """" Then
        startPos = startPos + 1
        endPos = InStr(startPos, json, """")
        ExtractJsonValueByKey = Mid$(json, startPos, endPos - startPos)
        Exit Function
    End If
    
    endPos = startPos
    Do While endPos <= Len(json) And Mid$(json, endPos, 1) Like "[0-9-]"
        endPos = endPos + 1
    Loop
    
    ExtractJsonValueByKey = Trim(Mid$(json, startPos, endPos - startPos))
End Function

Private Function ExtractJsonArray(ByVal json As String, key As String) As String

    Dim startPos As Long, endPos As Long, depth As Long
    
    startPos = InStr(json, """" & key & """:")
    If startPos = 0 Then Exit Function
    
    startPos = InStr(startPos, json, "[")
    If startPos = 0 Then Exit Function
    
    depth = 1
    endPos = startPos + 1
    
    Do While endPos <= Len(json) And depth > 0
        Select Case Mid$(json, endPos, 1)
            Case "[": depth = depth + 1
            Case "]": depth = depth - 1
        End Select
        endPos = endPos + 1
    Loop
    
    ExtractJsonArray = Mid$(json, startPos + 1, endPos - startPos - 2)
End Function

Private Function SplitNodes(nodesJson As String) As Variant
    Dim items As Collection
    Dim result() As String
    Dim i As Long, startPos As Long, depth As Long
    Dim ch As String
    
    Set items = New Collection
    startPos = 1
    depth = 0
    
    For i = 1 To Len(nodesJson)
        ch = Mid$(nodesJson, i, 1)
        
        If ch = "{" Then
            If depth = 0 Then startPos = i
            depth = depth + 1
        ElseIf ch = "}" Then
            depth = depth - 1
            If depth = 0 Then
                items.Add Mid$(nodesJson, startPos, i - startPos + 1)
            End If
        End If
    Next i
    
    ReDim result(items.count - 1)
    For i = 1 To items.count
        result(i - 1) = items(i)
    Next i
    
    SplitNodes = result
End Function


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
    Dim listIdx As Long

    lstPyramid.Clear

    For i = 0 To nodeCount - 1
        indent = String(nodes(i).depth * 2, " ")
        If nodes(i).depth > 0 Then indent = indent & "- "

        displayText = indent & nodes(i).text

        lstPyramid.AddItem displayText
        lstPyramid.List(lstPyramid.ListCount - 1, 1) = CStr(i)   ' store true node index
    Next i

    If selectedIndex >= 0 Then
        listIdx = FindListIndexForNode(selectedIndex)
        If listIdx >= 0 Then lstPyramid.ListIndex = listIdx
    End If
End Sub




Private Sub LoadPyramidFromTags()
    Dim structureData As String
    Dim chunk As String
    Dim i As Long
    Dim items() As String
    Dim parts() As String
    
    txtSituation.text = ActivePresentation.Tags("InstrumentaPyramidSCQ_Situation")
    txtComplication.text = ActivePresentation.Tags("InstrumentaPyramidSCQ_Complication")
    txtQuestion.text = ActivePresentation.Tags("InstrumentaPyramidSCQ_Question")
    
    structureData = ""
    i = 1
    
    Do
        chunk = ActivePresentation.Tags("InstrumentaPyramidSCQ_Structure_" & i)
        If chunk = "" Then Exit Do
        structureData = structureData & chunk
        i = i + 1
    Loop
    
   
    If structureData = "" Then Exit Sub
    
    items = Split(structureData, ";")
    nodeCount = 0
    ReDim nodes(UBound(items))
    
    For i = 0 To UBound(items)
        parts = Split(items(i), "|")
        If UBound(parts) >= 2 Then
            nodes(nodeCount).text = UnescapeDelimiter(parts(0))
            nodes(nodeCount).depth = CLng(parts(1))
            nodes(nodeCount).parentIndex = CLng(parts(2))
            nodeCount = nodeCount + 1
        End If
    Next i
End Sub


Private Sub SavePyramidToTags()
    Dim structureData As String
    Dim i As Long
    Dim chunkSize As Long: chunkSize = 1800
    Dim chunkCount As Long
    Dim pos As Long
    Dim chunk As String
    
    ActivePresentation.Tags.Add "InstrumentaPyramidSCQ_Situation", Trim(txtSituation.text)
    ActivePresentation.Tags.Add "InstrumentaPyramidSCQ_Complication", Trim(txtComplication.text)
    ActivePresentation.Tags.Add "InstrumentaPyramidSCQ_Question", Trim(txtQuestion.text)
    
    structureData = ""
    For i = 0 To nodeCount - 1
        If i > 0 Then structureData = structureData & ";"
        'structureData = structureData & nodes(i).text & "|" & nodes(i).depth & "|" & nodes(i).parentIndex
        structureData = structureData & EscapeDelimiter(nodes(i).text) & "|" & nodes(i).depth & "|" & nodes(i).parentIndex
    Next i
    
    i = 1
    Do While ActivePresentation.Tags("InstrumentaPyramidSCQ_Structure_" & i) <> ""
        ActivePresentation.Tags.Delete "InstrumentaPyramidSCQ_Structure_" & i
        i = i + 1
    Loop
    
    pos = 1
    chunkCount = 1
    
    Do While pos <= Len(structureData)
        chunk = Mid$(structureData, pos, chunkSize)
        ActivePresentation.Tags.Add "InstrumentaPyramidSCQ_Structure_" & chunkCount, chunk
        pos = pos + chunkSize
        chunkCount = chunkCount + 1
    Loop
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
    targetPosition = pres.Slides.count + 1
    
    slideIdx = slideIdx + 1
    Set sld = FindPyramidSlide(pres, slideIdx)
    
    If sld Is Nothing Then
        Set sld = pres.Slides.Add(targetPosition, ppLayoutText)
        sld.Tags.Add "InstrumentaPyramidSlideIndex", CStr(slideIdx)
        targetPosition = targetPosition + 1
    Else
        Call MoveSlideToPosition(sld, slideIdx, pres)
    End If
    
    sld.shapes.title.TextFrame.textRange.text = "Situation - Complication - Question"
    
    On Error Resume Next
    sld.shapes.Placeholders(2).Delete
    On Error GoTo 0
    
    colWidth = (pres.PageSetup.slideWidth - 150) / 3
    
    Set shp = Nothing
    For Each shp In sld.shapes
        If shp.Tags("InstrumentaPyramidElement") = "Situation" Then Exit For
        Set shp = Nothing
    Next shp
    If shp Is Nothing Then
        Set shp = sld.shapes.AddTextbox(msoTextOrientationHorizontal, 50, 120, colWidth, 300)
        shp.Tags.Add "InstrumentaPyramidElement", "Situation"
    End If
    With shp
        .TextFrame2.WordWrap = msoTrue
        .TextFrame2.textRange.text = "Situation" & vbCrLf & vbCrLf & txtSituation.text
        .TextFrame2.textRange.Font.Size = 14
        .TextFrame2.textRange.Paragraphs(1).Font.Bold = msoTrue
        .TextFrame2.textRange.Paragraphs(1).Font.Size = 18
    End With
    
    Set shp = Nothing
    For Each shp In sld.shapes
        If shp.Tags("InstrumentaPyramidElement") = "Complication" Then Exit For
        Set shp = Nothing
    Next shp
    leftPos = 50 + colWidth + 25
    If shp Is Nothing Then
        Set shp = sld.shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, 120, colWidth, 300)
        shp.Tags.Add "InstrumentaPyramidElement", "Complication"
    End If
    With shp
        .TextFrame2.WordWrap = msoTrue
        .TextFrame2.textRange.text = "Complication" & vbCrLf & vbCrLf & txtComplication.text
        .TextFrame2.textRange.Font.Size = 14
        .TextFrame2.textRange.Paragraphs(1).Font.Bold = msoTrue
        .TextFrame2.textRange.Paragraphs(1).Font.Size = 18
    End With
    
    Set shp = Nothing
    For Each shp In sld.shapes
        If shp.Tags("InstrumentaPyramidElement") = "Question" Then Exit For
        Set shp = Nothing
    Next shp
    leftPos = leftPos + colWidth + 25
    If shp Is Nothing Then
        Set shp = sld.shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, 120, colWidth, 300)
        shp.Tags.Add "InstrumentaPyramidElement", "Question"
    End If
    With shp
        .TextFrame2.WordWrap = msoTrue
        .TextFrame2.textRange.text = "Question" & vbCrLf & vbCrLf & txtQuestion.text
        .TextFrame2.textRange.Font.Size = 14
        .TextFrame2.textRange.Paragraphs(1).Font.Bold = msoTrue
        .TextFrame2.textRange.Paragraphs(1).Font.Size = 18
    End With
    
    slideIdx = slideIdx + 1
    Set sld = FindPyramidSlide(pres, slideIdx)
    
    If sld Is Nothing Then
        Set sld = pres.Slides.Add(targetPosition, ppLayoutText)
        sld.Tags.Add "InstrumentaPyramidSlideIndex", CStr(slideIdx)
        targetPosition = targetPosition + 1
    Else
        Call MoveSlideToPosition(sld, slideIdx, pres)
    End If
    
    sld.shapes.title.TextFrame.textRange.text = "Management Summary"
    
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
    
    Set tr = sld.shapes.Placeholders(2).TextFrame2.textRange
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
        
        sld.shapes.title.TextFrame.textRange.text = nodes(i).text
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
    firstPyramidPos = 0
    
    For Each s In pres.Slides
        If s.Tags("InstrumentaPyramidSlideIndex") = "1" Then
            firstPyramidPos = s.slideIndex
            Exit For
        End If
    Next s
    

    If firstPyramidPos = 0 Then

        targetPos = pres.Slides.count
    Else

        targetPos = firstPyramidPos + pyramidSlidesBefore
    End If
    

    If targetPos < 1 Then targetPos = 1
    If targetPos > pres.Slides.count Then targetPos = pres.Slides.count
    

    If currentPos <> targetPos And targetPos >= 1 And targetPos <= pres.Slides.count Then
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
    
    If paraNum < 1 Then
        GetNodeIndexForParagraph = -1
        Exit Function
    End If
    
    count = 0
    For i = 1 To nodeCount - 1
        If i > UBound(nodes) Then Exit For
        
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
    Dim i As Long
    Dim tagName As String
    
    Set pres = ActivePresentation

    On Error Resume Next
    pres.Tags.Delete "InstrumentaPyramidSCQ_Situation"
    pres.Tags.Delete "InstrumentaPyramidSCQ_Complication"
    pres.Tags.Delete "InstrumentaPyramidSCQ_Question"
    On Error GoTo 0

    i = 1
    Do
        tagName = "InstrumentaPyramidSCQ_Structure_" & i
        If pres.Tags(tagName) = "" Then Exit Do
        pres.Tags.Delete tagName
        i = i + 1
    Loop

    For Each sld In pres.Slides
        On Error Resume Next
        sld.Tags.Delete "InstrumentaPyramidSlideIndex"
        On Error GoTo 0
        
        For Each shp In sld.shapes
            On Error Resume Next
            shp.Tags.Delete "InstrumentaPyramidElement"
            On Error GoTo 0
        Next shp
    Next sld

    MsgBox "All Instrumenta Pyramid tags have been removed.", _
           vbInformation, "Cleanup Complete"

End Sub

Private Function EscapeDelimiter(text As String) As String
    EscapeDelimiter = Replace(text, "|", "~~PIPE~~")
End Function

Private Function UnescapeDelimiter(text As String) As String
    UnescapeDelimiter = Replace(text, "~~PIPE~~", "|")
End Function

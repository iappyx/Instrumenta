Attribute VB_Name = "ModuleTableOptimizeHeight"
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


Sub OptimizeTableHeightQuick()
    Call OptimizeTableHeightByContent(0)
End Sub

Sub OptimizeTableHeight3Iterations()
    Call OptimizeTableHeightByContent(3)
End Sub

Sub OptimizeTableHeight5Iterations()
    Call OptimizeTableHeightByContent(5)
End Sub

Sub OptimizeTableHeight10Iterations()
    Call OptimizeTableHeightByContent(10)
End Sub

Sub OptimizeTableHeight20Iterations()
    Call OptimizeTableHeightByContent(20)
End Sub

Sub OptimizeTableHeightByContent(numRuns As Integer)
    Dim optimizeTableShape As shape
    Dim optimizeTable As table
    Dim colIndex    As Integer, rowIndex As Integer
    Dim totalWidth  As Single
    Dim textLength() As Integer
    Dim sumTextLength As Integer
    Dim colWidths() As Single
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
        Exit Sub
    End If
    
    Set optimizeTableShape = ActiveWindow.Selection.ShapeRange(1)
    
    If optimizeTableShape.HasTable Then
        Set optimizeTable = optimizeTableShape.table
    End If
    
    If optimizeTable Is Nothing Then Exit Sub
    
    totalWidth = optimizeTableShape.Width
    
    ReDim textLength(1 To optimizeTable.Columns.Count)
    ReDim colWidths(1 To optimizeTable.Columns.Count)
    
    For colIndex = 1 To optimizeTable.Columns.Count
        textLength(colIndex) = 0
        For rowIndex = 1 To optimizeTable.Rows.Count
            textLength(colIndex) = textLength(colIndex) + Len(optimizeTable.Cell(rowIndex, colIndex).shape.TextFrame.textRange.Text)
        Next
        sumTextLength = sumTextLength + textLength(colIndex)
    Next
    
    For colIndex = 1 To optimizeTable.Columns.Count
        colWidths(colIndex) = (textLength(colIndex) / sumTextLength) * totalWidth
        optimizeTable.Columns(colIndex).Width = colWidths(colIndex)
    Next
    
    If numRuns > 0 Then
        Call OptimizeTableUsingCellTestingMultipleRuns(numRuns)
    End If
    
End Sub

Sub OptimizeTableUsingCellTestingMultipleRuns(numRuns As Integer)
    Dim optimizeTableShape As shape
    Dim originalTable As table
    Dim optimizeTable As table
    Dim colIndex As Integer, rowIndex As Integer
    Dim totalWidth As Single
    Dim stepSize As Single
    Dim increment As Integer
    Dim maxIncrements As Integer
    Dim currentTableHeight As Single
    Dim testResults() As Variant
    Dim bestWidths() As Single
    Dim minHeight As Single
    Dim totalAdjustedWidth As Single
    Dim runIndex As Integer
    Dim lastHeights(1 To 5) As Single
    Dim globalBestWidths() As Single
    Dim globalMinHeight As Single

    Set MyDocument = Application.ActiveWindow

    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
        Exit Sub
    End If
    
    Set optimizeTableShape = ActiveWindow.Selection.ShapeRange(1)
    If optimizeTableShape.HasTable Then
        Set originalTable = optimizeTableShape.table
    Else
        Exit Sub
    End If

    stepSize = Round(CalculateAverageFontSizeByParagraph(optimizeTableShape) / 2)
    maxIncrements = 5
    totalWidth = optimizeTableShape.Width
    ReDim bestWidths(1 To originalTable.Columns.Count)
    ReDim globalBestWidths(1 To originalTable.Columns.Count)
    globalMinHeight = 1E+30

    ProgressForm.Show

    For runIndex = 1 To numRuns
        SetProgress (runIndex / numRuns * 100)

        ReDim testResults(1 To originalTable.Rows.Count * originalTable.Columns.Count * maxIncrements, 1 To 4 + originalTable.Columns.Count)
        Dim resultIndex As Integer
        resultIndex = 1

        For colIndex = 1 To originalTable.Columns.Count
            SetProgress (runIndex / numRuns * 100), "Iteration: " & runIndex & " of " & numRuns & ", column " & colIndex & " : Testing individual cells"
            For rowIndex = 1 To originalTable.Rows.Count
                For increment = 1 To maxIncrements
                    Set duplicateTableShape = optimizeTableShape.Duplicate
                    Set optimizeTable = duplicateTableShape.table

                    Dim originalWidth As Single
                    originalWidth = optimizeTable.Columns(colIndex).Width
                    optimizeTable.Columns(colIndex).Width = originalWidth + (increment * stepSize)

                    totalAdjustedWidth = (increment * stepSize) / (optimizeTable.Columns.Count - 1)
                    Dim otherColIndex As Integer
                    For otherColIndex = 1 To optimizeTable.Columns.Count
                        If otherColIndex <> colIndex Then
                            optimizeTable.Columns(otherColIndex).Width = optimizeTable.Columns(otherColIndex).Width - totalAdjustedWidth
                        End If
                    Next

                    currentTableHeight = duplicateTableShape.Height

                    testResults(resultIndex, 1) = rowIndex
                    testResults(resultIndex, 2) = colIndex
                    testResults(resultIndex, 3) = increment
                    testResults(resultIndex, 4) = currentTableHeight
                    Dim colWidthsArray() As Single
                    ReDim colWidthsArray(1 To optimizeTable.Columns.Count)
                    For otherColIndex = 1 To optimizeTable.Columns.Count
                        testResults(resultIndex, 4 + otherColIndex) = optimizeTable.Columns(otherColIndex).Width
                    Next
                    resultIndex = resultIndex + 1

                    duplicateTableShape.Delete
                Next
            Next
        Next

        minHeight = 1E+30
        Dim testRow As Integer
        SetProgress (runIndex / numRuns * 100), "Iteration: " & runIndex & " of " & numRuns & ": Analyzing results"
        For testRow = 1 To resultIndex - 1
            If testResults(testRow, 4) < minHeight Then
                minHeight = testResults(testRow, 4)
                For colIndex = 1 To originalTable.Columns.Count
                    bestWidths(colIndex) = testResults(testRow, 4 + colIndex)
                Next
            End If
        Next

        If minHeight < globalMinHeight Then
            globalMinHeight = minHeight
            For colIndex = 1 To originalTable.Columns.Count
                globalBestWidths(colIndex) = bestWidths(colIndex)
            Next
        End If

        For colIndex = 1 To originalTable.Columns.Count
            originalTable.Columns(colIndex).Width = bestWidths(colIndex)
        Next

        lastHeights((runIndex - 1) Mod 5 + 1) = minHeight
       ' If runIndex >= 5 Then
       '     If lastHeights(1) = lastHeights(2) And lastHeights(2) = lastHeights(3) And lastHeights(3) = lastHeights(4) And lastHeights(4) = lastHeights(5) Then
       '         Exit For
       '     End If
       ' End If
    Next

    For colIndex = 1 To originalTable.Columns.Count
        originalTable.Columns(colIndex).Width = globalBestWidths(colIndex)
    Next

    ProgressForm.Hide
    Unload ProgressForm
    Set optimizeTableShape = Nothing
    Set originalTable = Nothing
    Set optimizeTable = Nothing
    Set duplicateTableShape = Nothing
    Erase testResults
    Erase bestWidths
    Erase globalBestWidths
End Sub

Sub OptimizeTableUsingCellTestingMultipleRunsOld(numRuns As Integer)
    'Old sub as backup
    Dim optimizeTableShape As shape
    Dim originalTable As table
    Dim optimizeTable As table
    Dim colIndex    As Integer, rowIndex As Integer
    Dim totalWidth  As Single
    Dim stepSize    As Single
    Dim increment   As Integer
    Dim maxIncrements As Integer
    Dim currentTableHeight As Single
    Dim testResults() As Variant
    Dim bestWidths() As Single
    Dim minHeight   As Single
    Dim totalAdjustedWidth As Single
    Dim runIndex    As Integer
    Dim lastHeights(1 To 5) As Single
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
        Exit Sub
    End If
    
    Set optimizeTableShape = ActiveWindow.Selection.ShapeRange(1)
    
    If optimizeTableShape.HasTable Then
        Set originalTable = optimizeTableShape.table
    Else
        
        Exit Sub
        
    End If
    
    stepSize = Round(CalculateAverageFontSizeByParagraph(optimizeTableShape) / 2)
    maxIncrements = 5
    
    totalWidth = optimizeTableShape.Width
    ReDim bestWidths(1 To originalTable.Columns.Count)
    
    ProgressForm.Show
    
    For runIndex = 1 To numRuns
        
        SetProgress (runIndex / numRuns * 100)
        
        ReDim testResults(1 To originalTable.Rows.Count * originalTable.Columns.Count * maxIncrements, 1 To 4 + originalTable.Columns.Count)
        
        Dim resultIndex As Integer
        resultIndex = 1
        
        For colIndex = 1 To originalTable.Columns.Count
            SetProgress (runIndex / numRuns * 100), "Iteration: " & runIndex & " of " & numRuns & ", column " & colIndex & " : Testing individual cells"
            For rowIndex = 1 To originalTable.Rows.Count
                
                For increment = 1 To maxIncrements
                    
                    Set duplicateTableShape = optimizeTableShape.Duplicate
                    Set optimizeTable = duplicateTableShape.table
                    
                    Dim originalWidth As Single
                    originalWidth = optimizeTable.Columns(colIndex).Width
                    optimizeTable.Columns(colIndex).Width = originalWidth + (increment * stepSize)
                    
                    totalAdjustedWidth = (increment * stepSize) / (optimizeTable.Columns.Count - 1)
                    Dim otherColIndex As Integer
                    For otherColIndex = 1 To optimizeTable.Columns.Count
                        If otherColIndex <> colIndex Then
                            optimizeTable.Columns(otherColIndex).Width = optimizeTable.Columns(otherColIndex).Width - totalAdjustedWidth
                        End If
                    Next
                    
                    currentTableHeight = duplicateTableShape.Height
                    
                    testResults(resultIndex, 1) = rowIndex
                    testResults(resultIndex, 2) = colIndex
                    testResults(resultIndex, 3) = increment
                    testResults(resultIndex, 4) = currentTableHeight
                    Dim colWidthsArray() As Single
                    ReDim colWidthsArray(1 To optimizeTable.Columns.Count)
                    For otherColIndex = 1 To optimizeTable.Columns.Count
                        testResults(resultIndex, 4 + otherColIndex) = optimizeTable.Columns(otherColIndex).Width
                    Next
                    resultIndex = resultIndex + 1
                    
                    duplicateTableShape.Delete
                Next
            Next
        Next
        
        minHeight = 1E+30
        
        Dim testRow As Integer
        SetProgress (runIndex / numRuns * 100), "Iteration: " & runIndex & " of " & numRuns & ": Analyzing results"
        For testRow = 1 To resultIndex - 1
            If testResults(testRow, 4) < minHeight Then
                minHeight = testResults(testRow, 4)
                For colIndex = 1 To originalTable.Columns.Count
                    bestWidths(colIndex) = testResults(testRow, 4 + colIndex)
                Next
            End If
        Next
        
        For colIndex = 1 To originalTable.Columns.Count
            originalTable.Columns(colIndex).Width = bestWidths(colIndex)
        Next
        
        lastHeights((runIndex - 1) Mod 5 + 1) = minHeight
        
       ' If runIndex >= 5 Then
       '     If lastHeights(1) = lastHeights(2) And lastHeights(2) = lastHeights(3) And lastHeights(3) = lastHeights(4) And lastHeights(4) = lastHeights(5) Then
       '         Exit For
       '     End If
       ' End If
        
    Next
    
    ProgressForm.Hide
    Unload ProgressForm
    
    Set optimizeTableShape = Nothing
    Set originalTable = Nothing
    Set optimizeTable = Nothing
    Set duplicateTableShape = Nothing
    
    Erase testResults
    Erase bestWidths
    
End Sub

Function GetColumnWidths(sourceTable As table) As Variant
    Dim colWidths() As Single
    Dim colIndex    As Integer
    ReDim colWidths(1 To sourceTable.Columns.Count)
    
    For colIndex = 1 To sourceTable.Columns.Count
        colWidths(colIndex) = sourceTable.Columns(colIndex).Width
    Next
    
    GetColumnWidths = colWidths
End Function

Function CalculateAverageFontSizeByParagraph(sourceTableShape As shape) As Single
    Dim sourceTable As table
    Dim colIndex    As Integer, rowIndex As Integer
    Dim fontSizeSum As Single
    Dim paragraphCount As Integer
    Dim averageFontSize As Single
    Dim textRange   As textRange
    Dim paragraphRange As textRange
    Dim paragraphIndex As Integer
    
    Set sourceTable = sourceTableShape.table
    
    fontSizeSum = 0
    paragraphCount = 0
    
    For colIndex = 1 To sourceTable.Columns.Count
        For rowIndex = 1 To sourceTable.Rows.Count
            
            Set textRange = sourceTable.Cell(rowIndex, colIndex).shape.TextFrame.textRange
            
            If Len(Trim(textRange.Text)) > 0 Then
                
                For paragraphIndex = 1 To textRange.Paragraphs.Count
                    Set paragraphRange = textRange.Paragraphs(paragraphIndex)
                    fontSizeSum = fontSizeSum + paragraphRange.Font.Size
                    paragraphCount = paragraphCount + 1
                Next
            End If
        Next
    Next
    
    If paragraphCount > 0 Then
        averageFontSize = fontSizeSum / paragraphCount
    Else
        averageFontSize = 0
    End If
    
    CalculateAverageFontSizeByParagraph = averageFontSize
End Function

Attribute VB_Name = "ModuleCallbacks"
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

Sub TableColumnGapsEven()
    TableColumnGaps "even", 5
End Sub

Sub TableColumnGapsOdd()
    TableColumnGaps "odd", 5
End Sub

Sub TableRowGapsEven()
    TableRowGaps "even", 5
End Sub

Sub TableRowGapsOdd()
    TableRowGaps "odd", 5
End Sub

Sub CopySlideNotesToWord()
    CopySlideNotesToClipboard True
End Sub

Sub CopySlideNotesToClipboardOnly()
    CopySlideNotesToClipboard False
End Sub

Sub CopyStorylineToWord()
    CopyStorylineToClipboard True
End Sub

Sub CopyStorylineToClipBoardOnly()
    CopyStorylineToClipboard False
End Sub

Sub ConnectRectangleShapesRightToLeft()
    ConnectRectangleShapes "RightToLeft"
End Sub

Sub ConnectRectangleShapesLeftToRight()
    ConnectRectangleShapes "LeftToRight"
End Sub

Sub ConnectRectangleShapesBottomToTop()
    ConnectRectangleShapes "BottomToTop"
End Sub

Sub ConnectRectangleShapesTopToBottom()
    ConnectRectangleShapes "TopToBottom"
End Sub

Sub TextInsertEuro()
    ObjectsTextInsertSpecialCharacter 8364
End Sub

Sub TextInsertCopyright()
    ObjectsTextInsertSpecialCharacter 169
End Sub

Sub GenerateStampConfidential()
    Dim StampColor As Long
    StampColor = RGB(192, 0, 0)
    GenerateStamp "CONFIDENTIAL", StampColor
End Sub

Sub GenerateStampDoNotDistribute()
    Dim StampColor As Long
    StampColor = RGB(192, 0, 0)
    GenerateStamp "DO NOT DISTRIBUTE", StampColor
End Sub

Sub GenerateStampDraft()
    Dim StampColor As Long
    StampColor = RGB(0, 112, 192)
    GenerateStamp "DRAFT", StampColor
End Sub

Sub GenerateStampUpdated()
    Dim StampColor As Long
    StampColor = RGB(255, 153, 0)
    GenerateStamp "UPDATED", StampColor
End Sub

Sub GenerateStampNew()
    Dim StampColor As Long
    StampColor = RGB(0, 176, 80)
    GenerateStamp "NEW", StampColor
End Sub

Sub GenerateStampToBeRemoved()
    Dim StampColor As Long
    StampColor = RGB(179, 0, 0)
    GenerateStamp "TO BE REMOVED", StampColor
End Sub

Sub GenerateStampToAppendix()
    Dim StampColor As Long
    StampColor = RGB(127, 127, 127)
    GenerateStamp "TO APPENDIX", StampColor
End Sub

Sub GenerateHarveyBallCustom()
    CustomPercentage = CInt(InputBox("Harvey ball percentage:", "Custom HarveyBall", 50))
    GenerateHarveyBallPercent (CustomPercentage)
End Sub

Sub GenerateHarveyBall25()
    GenerateHarveyBallPercent (25)
End Sub

Sub GenerateHarveyBall33()
    GenerateHarveyBallPercent (33)
End Sub

Sub GenerateHarveyBall50()
    GenerateHarveyBallPercent (50)
End Sub

Sub GenerateHarveyBall67()
    GenerateHarveyBallPercent (67)
End Sub

Sub GenerateHarveyBall75()
    GenerateHarveyBallPercent (75)
End Sub

Sub GenerateHarveyBall100()
    GenerateHarveyBallPercent (100)
End Sub

Sub GenerateHarveyBall10()
    GenerateHarveyBallPercent (10)
End Sub

Sub GenerateHarveyBall20()
    GenerateHarveyBallPercent (20)
End Sub

Sub GenerateHarveyBall30()
    GenerateHarveyBallPercent (30)
End Sub

Sub GenerateHarveyBall40()
    GenerateHarveyBallPercent (40)
End Sub

Sub GenerateHarveyBall60()
    GenerateHarveyBallPercent (60)
End Sub

Sub GenerateHarveyBall70()
    GenerateHarveyBallPercent (70)
End Sub

Sub GenerateHarveyBall80()
    GenerateHarveyBallPercent (80)
End Sub

Sub GenerateHarveyBall90()
    GenerateHarveyBallPercent (90)
End Sub

Sub GenerateFivePointStars05()
    GenerateFivePointStars (0.5)
End Sub

Sub GenerateFivePointStars10()
    GenerateFivePointStars (1)
End Sub

Sub GenerateFivePointStars15()
    GenerateFivePointStars (1.5)
End Sub

Sub GenerateFivePointStars20()
    GenerateFivePointStars (2)
End Sub

Sub GenerateFivePointStars25()
    GenerateFivePointStars (2.5)
End Sub

Sub GenerateFivePointStars30()
    GenerateFivePointStars (3)
End Sub

Sub GenerateFivePointStars35()
    GenerateFivePointStars (3.5)
End Sub

Sub GenerateFivePointStars40()
    GenerateFivePointStars (4)
End Sub

Sub GenerateFivePointStars45()
    GenerateFivePointStars (4.5)
End Sub

Sub GenerateFivePointStars50()
    GenerateFivePointStars (5)
End Sub

Sub GenerateRAGStatusRed()
    GenerateRAGStatus ("red")
End Sub

Sub GenerateRAGStatusAmber()
    GenerateRAGStatus ("amber")
End Sub

Sub GenerateRAGStatusGreen()
    GenerateRAGStatus ("green")
End Sub

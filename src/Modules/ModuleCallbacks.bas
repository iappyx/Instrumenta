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


Public InstrumentaRibbon As IRibbonUI
Public InstrumentaVisible As String
Dim SetPositionAppEventHandler As New SetPositionEventsClass

Sub InitialiseSetPositionAppEventHandler()

    Set SetPositionAppEventHandler.App = Application
    SetPositionForm.Show 0
      
End Sub

Sub UnloadSetPositionAppEventHandler()

    Set SetPositionAppEventHandler.App = Nothing
    Unload SetPositionForm
    
End Sub

 
Sub InstrumentaInitialize(Ribbon As IRibbonUI)
    InstrumentaVisible = "InstrumentaVisible"
    Set InstrumentaRibbon = Ribbon
    InitializeEmojis
    InitializeEmojis2 'for 32-bit compatibility split between 2
    InitializeEmojiNames
    InstrumentaRibbon.Invalidate
    'InstrumentaRibbon.ActivateTab "InstrumentaPowerpointToolbar"
    
    'set Pro or Review mode
    If GetSetting("Instrumenta", "General", "OperatingMode", "pro") = "review" Then
     Call InstrumentaRefresh(UpdateTag:="*R*")
    End If
    
End Sub

Sub InstrumentaRefresh(UpdateTag As String)

    InstrumentaVisible = UpdateTag
    If Not InstrumentaRibbon Is Nothing Then
    
    InstrumentaRibbon.Invalidate
       
    End If

End Sub


Sub InstrumentaGetVisible(control As IRibbonControl, ByRef visible)
    If InstrumentaVisible = "InstrumentaVisible" Then
        visible = True
    Else
        If control.Tag Like InstrumentaVisible Then
            visible = True
        Else
            visible = False
        End If
    End If
End Sub


Sub EmojiGallery_GetItemImage(control As IRibbonControl, index As Integer, ByRef returnedVal)

#If Mac Then
'Even though there is no picture needed, you have to set it for Mac compatibility
returnedVal = "AppointmentColor4"
#End If

End Sub

Sub RibbonObjectGetImage(control As IRibbonControl, ByRef returnedVal)


#If Mac Then
    
    Select Case control.id
        Case "FivePointStarMenu"
            returnedVal = "ShapeStar"
        Case "StampsMenu"
            returnedVal = "CustomHeaderGallery"
        Case "ObjectsSizeToTallest"
            returnedVal = "ObjectNudgeDown"
        Case "ObjectsSizeShortest"
            returnedVal = "ObjectNudgeUp"
        Case "ObjectsSizeToWidest"
            returnedVal = "ObjectNudgeRight"
        Case "ObjectsSizeNarrowest"
            returnedVal = "ObjectNudgeLeft"
    End Select
    
#Else
    
    Select Case control.id
        Case "FivePointStarMenu"
            returnedVal = "StarRatedFull"
        Case "StampsMenu"
            returnedVal = "StampTool"
        Case "ObjectsSizeToTallest"
            returnedVal = "SizeToTallest"
        Case "ObjectsSizeShortest"
            returnedVal = "SizeToShortest"
        Case "ObjectsSizeToWidest"
            returnedVal = "SizeToWidest"
        Case "ObjectsSizeNarrowest"
            returnedVal = "SizeToNarrowest"
    End Select

#End If

End Sub



Sub EmojiGallery_GetItemWidth(control As IRibbonControl, ByRef returnedVal)

#If Mac Then
'Even though there is no picture needed, you have to set it for Mac compatibility
returnedVal = 25
#End If


End Sub

Sub EmojiGallery_GetItemHeight(control As IRibbonControl, ByRef returnedVal)

#If Mac Then
'Even though there is no picture needed, you have to set it for Mac compatibility
returnedVal = 25
#End If


End Sub


Sub EmojiGallery_GetItemCount(control As IRibbonControl, ByRef returnedVal)
    
    Select Case control.id
     Case "EmojiGallery1"
        returnedVal = 111
     Case "EmojiGallery2"
        returnedVal = 49
     Case "EmojiGallery3"
        returnedVal = 104
     Case "EmojiGallery4"
        returnedVal = 48
     Case "EmojiGallery5"
        returnedVal = 182
     Case "EmojiGallery6"
        returnedVal = 122
     Case "EmojiGallery7"
        returnedVal = 80
     Case "EmojiGallery8"
        returnedVal = 117
     Case "EmojiGallery9"
        returnedVal = 196
     Case "EmojiGallery10"
        returnedVal = 175
    End Select
    
End Sub

Sub EmojiGallery_GetItemID(control As IRibbonControl, index As Integer, ByRef returnedVal)
    
    Select Case control.id
     Case "EmojiGallery1"
        returnedVal = "Emoji" & index
     Case "EmojiGallery2"
        returnedVal = "Emoji" & index + 111
     Case "EmojiGallery3"
        returnedVal = "Emoji" & index + 111 + 49
     Case "EmojiGallery4"
        returnedVal = "Emoji" & index + 111 + 49 + 104
     Case "EmojiGallery5"
        returnedVal = "Emoji" & index + 111 + 49 + 104 + 48
     Case "EmojiGallery6"
        returnedVal = "Emoji" & index + 111 + 49 + 104 + 48 + 182
     Case "EmojiGallery7"
        returnedVal = "Emoji" & index + 111 + 49 + 104 + 48 + 182 + 122
     Case "EmojiGallery8"
        returnedVal = "Emoji" & index + 111 + 49 + 104 + 48 + 182 + 122 + 80
     Case "EmojiGallery9"
        returnedVal = "Emoji" & index + 111 + 49 + 104 + 48 + 182 + 122 + 80 + 117
     Case "EmojiGallery10"
        returnedVal = "Emoji" & index + 111 + 49 + 104 + 48 + 182 + 122 + 80 + 117 + 196
    End Select
    
End Sub

Sub EmojiGallery_GetItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)

    Select Case control.id
     Case "EmojiGallery1"
        returnedVal = AllEmojis(index + 1)
     Case "EmojiGallery2"
        returnedVal = AllEmojis(index + 1 + 111)
     Case "EmojiGallery3"
        returnedVal = AllEmojis(index + 1 + 111 + 49)
     Case "EmojiGallery4"
        returnedVal = AllEmojis(index + 1 + 111 + 49 + 104)
     Case "EmojiGallery5"
        returnedVal = AllEmojis(index + 1 + 111 + 49 + 104 + 48)
     Case "EmojiGallery6"
        returnedVal = AllEmojis(index + 1 + 111 + 49 + 104 + 48 + 182)
     Case "EmojiGallery7"
        returnedVal = AllEmojis(index + 1 + 111 + 49 + 104 + 48 + 182 + 122)
     Case "EmojiGallery8"
        returnedVal = AllEmojis(index + 1 + 111 + 49 + 104 + 48 + 182 + 122 + 80)
     Case "EmojiGallery9"
        returnedVal = AllEmojis(index + 1 + 111 + 49 + 104 + 48 + 182 + 122 + 80 + 117)
     Case "EmojiGallery10"
        returnedVal = AllEmojis(index + 1 + 111 + 49 + 104 + 48 + 182 + 122 + 80 + 117 + 196)
    End Select

End Sub


Sub EmojiGallery_GetItemScreentip(control As IRibbonControl, index As Integer, ByRef returnedVal)

    Select Case control.id
     Case "EmojiGallery1"
        returnedVal = AllEmojis(index + 1) & " " & StrConv(EmojiNames(index + 1), vbProperCase)
     Case "EmojiGallery2"
        returnedVal = AllEmojis(index + 1 + 111) & " " & StrConv(EmojiNames(index + 1 + 111), vbProperCase)
     Case "EmojiGallery3"
        returnedVal = AllEmojis(index + 1 + 111 + 49) & " " & StrConv(EmojiNames(index + 1 + 111 + 49), vbProperCase)
     Case "EmojiGallery4"
        returnedVal = AllEmojis(index + 1 + 111 + 49 + 104) & " " & StrConv(EmojiNames(index + 1 + 111 + 49 + 104), vbProperCase)
     Case "EmojiGallery5"
        returnedVal = AllEmojis(index + 1 + 111 + 49 + 104 + 48) & " " & StrConv(EmojiNames(index + 1 + 111 + 49 + 104 + 48), vbProperCase)
     Case "EmojiGallery6"
        returnedVal = AllEmojis(index + 1 + 111 + 49 + 104 + 48 + 182) & " " & StrConv(EmojiNames(index + 1 + 111 + 49 + 104 + 48 + 182), vbProperCase)
     Case "EmojiGallery7"
        returnedVal = AllEmojis(index + 1 + 111 + 49 + 104 + 48 + 182 + 122) & " " & StrConv(EmojiNames(index + 1 + 111 + 49 + 104 + 48 + 182 + 122), vbProperCase)
     Case "EmojiGallery8"
        returnedVal = AllEmojis(index + 1 + 111 + 49 + 104 + 48 + 182 + 122 + 80) & " " & StrConv(EmojiNames(index + 1 + 111 + 49 + 104 + 48 + 182 + 122 + 80), vbProperCase)
     Case "EmojiGallery9"
        returnedVal = AllEmojis(index + 1 + 111 + 49 + 104 + 48 + 182 + 122 + 80 + 117) & " " & StrConv(EmojiNames(index + 1 + 111 + 49 + 104 + 48 + 182 + 122 + 80 + 117), vbProperCase)
     Case "EmojiGallery10"
        returnedVal = AllEmojis(index + 1 + 111 + 49 + 104 + 48 + 182 + 122 + 80 + 117 + 196) & " " & StrConv(EmojiNames(index + 1 + 111 + 49 + 104 + 48 + 182 + 122 + 80 + 117 + 196), vbProperCase)
    End Select

End Sub


Sub EmojiGallery_OnAction(control As IRibbonControl, id As String, index As Integer)
    
    Select Case control.id
     Case "EmojiGallery1"
        GenerateEmoji (index + 1)
     Case "EmojiGallery2"
        GenerateEmoji (index + 1 + 111)
     Case "EmojiGallery3"
        GenerateEmoji (index + 1 + 111 + 49)
     Case "EmojiGallery4"
        GenerateEmoji (index + 1 + 111 + 49 + 104)
     Case "EmojiGallery5"
        GenerateEmoji (index + 1 + 111 + 49 + 104 + 48)
     Case "EmojiGallery6"
        GenerateEmoji (index + 1 + 111 + 49 + 104 + 48 + 182)
     Case "EmojiGallery7"
        GenerateEmoji (index + 1 + 111 + 49 + 104 + 48 + 182 + 122)
     Case "EmojiGallery8"
        GenerateEmoji (index + 1 + 111 + 49 + 104 + 48 + 182 + 122 + 80)
     Case "EmojiGallery9"
        GenerateEmoji (index + 1 + 111 + 49 + 104 + 48 + 182 + 122 + 80 + 117)
     Case "EmojiGallery10"
        GenerateEmoji (index + 1 + 111 + 49 + 104 + 48 + 182 + 122 + 80 + 117 + 196)
    End Select
    
End Sub

Sub ShowSlideLibrary()
    InsertSlideLibrarySlide.Show
End Sub

Sub ShowSettings()
    SettingsForm.Show
End Sub

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

Sub TextInsertNoBreakSpace()
    ObjectsTextInsertSpecialCharacter 160
End Sub


Sub TextInsertCopyright()
    ObjectsTextInsertSpecialCharacter 169
End Sub

Sub GenerateStampConfidential()
    Dim StampColor As Long
    StampColor = GetSetting("Instrumenta", "Stamps", "ConfidentialColor", "192")
    GenerateStamp "CONFIDENTIAL", StampColor
End Sub

Sub GenerateStampDoNotDistribute()
    Dim StampColor As Long
    StampColor = GetSetting("Instrumenta", "Stamps", "DoNotDistributeColor", "192")
    GenerateStamp "DO NOT DISTRIBUTE", StampColor
End Sub

Sub GenerateStampDraft()
    Dim StampColor As Long
    StampColor = GetSetting("Instrumenta", "Stamps", "DraftColor", "12611584")
    GenerateStamp "DRAFT", StampColor
End Sub

Sub GenerateStampUpdated()
    Dim StampColor As Long
    StampColor = GetSetting("Instrumenta", "Stamps", "UpdatedColor", "39423")
    GenerateStamp "UPDATED", StampColor
End Sub

Sub GenerateStampNew()
    Dim StampColor As Long
    StampColor = GetSetting("Instrumenta", "Stamps", "NewColor", "5287936")
    GenerateStamp "NEW", StampColor
End Sub

Sub GenerateStampToBeRemoved()
    Dim StampColor As Long
    StampColor = GetSetting("Instrumenta", "Stamps", "ToBeRemovedColor", "179")
    GenerateStamp "TO BE REMOVED", StampColor
End Sub

Sub GenerateStampToAppendix()
    Dim StampColor As Long
    StampColor = GetSetting("Instrumenta", "Stamps", "ToAppendixColor", "8355711")
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

Sub ColorBoldTextColorPicker()
ObjectsTextColorBold (False)
End Sub

Sub ColorBoldTextColorAutomatically()
ObjectsTextColorBold (True)
End Sub

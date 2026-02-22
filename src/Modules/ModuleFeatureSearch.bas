Attribute VB_Name = "ModuleFeatureSearch"
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

Public Type FeatureData
    id As String
    label As String
    OnAction As String
    GroupLabel As String
    TabSingleView As String
    TabMultiView As String
End Type

Private Features() As FeatureData
Private FeatureCount As Long

Public Sub ShowFeatureSearch()
    If FeatureCount = 0 Then
        LoadInstrumentaFeatures
    End If
    
    If FeatureCount > 0 Then
        FeatureSearchForm.Show 0
    Else
        MsgBox "No features found.", vbExclamation
    End If
End Sub

Private Sub LoadInstrumentaFeatures()
    On Error GoTo ErrorHandler
    
    ReDim Features(1 To 500) As FeatureData
    FeatureCount = 0
    
    AddFeature "BoldColorButton", "Color bold text (automatically)", "ColorBoldTextColorAutomatically", "Font", "Instrumenta > Font > Inside splitbutton 'Bold'", "Instrumenta [Text] > Font > Inside splitbutton 'Bold'"
    AddFeature "BoldColorPickerButton", "Color bold text (color picker)", "ColorBoldTextColorPicker", "Font", "Instrumenta > Font > Inside splitbutton 'Bold'", "Instrumenta [Text] > Font > Inside splitbutton 'Bold'"
    AddFeature "ObjectsTextDeleteStrikethroughButton", "Delete strikethrough text", "ObjectsTextDeleteStrikethrough", "Font", "Instrumenta > Font > Inside splitbutton 'Strikethrough'", "Instrumenta [Text] > Font > Inside splitbutton 'Strikethrough'"
    AddFeature "ObjectsToggleAutoSize", "Toggle autofit", "ObjectsToggleAutoSize", "Text", "Instrumenta > Text > Inside splitbutton 'Toggle autofit'", "Instrumenta [Text] > Text > Inside splitbutton 'Toggle autofit'"
    AddFeature "ObjectsAutoSizeNone", "Do not Autofit", "ObjectsAutoSizeNone", "Text", "Instrumenta > Text > Inside splitbutton 'Toggle autofit'", "Instrumenta [Text] > Text > Inside splitbutton 'Toggle autofit'"
    AddFeature "ObjectsAutoSizeTextToFitShape", "Resize text on overflow", "ObjectsAutoSizeTextToFitShape", "Text", "Instrumenta > Text > Inside splitbutton 'Toggle autofit'", "Instrumenta [Text] > Text > Inside splitbutton 'Toggle autofit'"
    AddFeature "ObjectsAutoSizeShapeToFitText", "Resize shape to fit text", "ObjectsAutoSizeShapeToFitText", "Text", "Instrumenta > Text > Inside splitbutton 'Toggle autofit'", "Instrumenta [Text] > Text > Inside splitbutton 'Toggle autofit'"
    AddFeature "BulletsTicks", "Ticks", "TextBulletsTicks", "Text", "Instrumenta > Text", "Instrumenta [Text] > Text"
    AddFeature "BulletsCrosses", "Crosses", "TextBulletsCrosses", "Text", "Instrumenta > Text", "Instrumenta [Text] > Text"
    AddFeature "TextInsertEuro", "Euro", "TextInsertEuro", "Text", "Instrumenta > Text > Inside menu 'Special characters'", "Instrumenta [Text] > Text > Inside menu 'Special characters'"
    AddFeature "TextInsertCopyright", "Copyright", "TextInsertCopyright", "Text", "Instrumenta > Text > Inside menu 'Special characters'", "Instrumenta [Text] > Text > Inside menu 'Special characters'"
    AddFeature "TextInsertNoBreakSpace", "' ' (NonBreakingSpace)", "TextInsertNoBreakSpace", "Text", "Instrumenta > Text > Inside menu 'Special characters'", "Instrumenta [Text] > Text > Inside menu 'Special characters'"
    AddFeature "IncreaseLineSpacing", "Increase line spacing", "ObjectsIncreaseLineSpacing", "Text", "Instrumenta > Text > Inside splitbutton 'Increase line spacing'", "Instrumenta [Text] > Text > Inside splitbutton 'Increase line spacing'"
    AddFeature "IncreaseLineSpacingBeforeAndAfter", "Increase line spacing before and after", "ObjectsIncreaseLineSpacingBeforeAndAfter", "Text", "Instrumenta > Text > Inside splitbutton 'Increase line spacing'", "Instrumenta [Text] > Text > Inside splitbutton 'Increase line spacing'"
    AddFeature "DecreaseLineSpacing", "Decrease line spacing", "ObjectsDecreaseLineSpacing", "Text", "Instrumenta > Text > Inside splitbutton 'Decrease line spacing'", "Instrumenta [Text] > Text > Inside splitbutton 'Decrease line spacing'"
    AddFeature "DecreaseLineSpacingBeforeAndAfter", "Decrease line spacing before and after", "ObjectsDecreaseLineSpacingBeforeAndAfter", "Text", "Instrumenta > Text > Inside splitbutton 'Decrease line spacing'", "Instrumenta [Text] > Text > Inside splitbutton 'Decrease line spacing'"
    AddFeature "ToggleTextWrap", "Toggle text wrap", "ObjectsTextWordwrapToggle", "Text", "Instrumenta > Text", "Instrumenta [Text] > Text"
    AddFeature "ObjectsTextSplitByParagraphButton", "Split (by paragraphs) into multiple shapes", "ObjectsTextSplitByParagraph", "Text", "Instrumenta > Text", "Instrumenta [Text] > Text"
    AddFeature "ObjectsTextMergeButton", "Merge text of all shapes in first selected shape", "ObjectsTextMerge", "Text", "Instrumenta > Text", "Instrumenta [Text] > Text"
    AddFeature "RemoveTextButton", "Remove text", "ObjectsRemoveText", "Text", "Instrumenta > Text > Inside splitbutton 'Remove text'", "Instrumenta [Text] > Text > Inside splitbutton 'Remove text'"
    AddFeature "RemoveHyperlinksButton", "Remove hyperlinks", "ObjectsRemoveHyperlinks", "Text", "Instrumenta > Text > Inside splitbutton 'Remove text'", "Instrumenta [Text] > Text > Inside splitbutton 'Remove text'"
    AddFeature "SwapTextButton", "Swap text (no formatting)", "ObjectsSwapTextNoFormatting", "Text", "Instrumenta > Text", "Instrumenta [Text] > Text"
    AddFeature "SwapTextButton2", "Swap text (with formatting)", "ObjectsSwapText", "Text", "Instrumenta > Text", "Instrumenta [Text] > Text"
    AddFeature "ChangeSpellCheckLanguage", "Set proofing language on all objects and all slides", "ShowChangeSpellCheckLanguageForm", "Text", "Instrumenta > Text", "Instrumenta [Text] > Text"
    AddFeature "ApplyH1", "Apply heading 1", "ApplyH1", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "ApplyH2", "Apply heading 2", "ApplyH2", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "ApplyH3", "Apply heading 3", "ApplyH3", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "ApplyParagraph", "Apply paragraph", "ApplyParagraph", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "ApplyQuote", "Apply quote", "ApplyQuote", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "ApplyCustom1", "Apply custom style 1", "ApplyCustom1", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "ApplyCustom2", "Apply custom style 2", "ApplyCustom2", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "ApplyCustom3", "Apply custom style 3", "ApplyCustom3", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "ApplyCustom4", "Apply custom style 4", "ApplyCustom4", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "ApplyCustom5", "Apply custom style 5", "ApplyCustom5", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "StylesUpdateFullShapeStyles", "Update shapes with styles applied to match stylesheet", "UpdateFullShapeStyles", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "StylesExportStylesToPPTX", "Export Stylesheet", "ExportStylesToPPTX", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "StylesImportStylesFromPPTX", "Import Stylesheet", "ImportStylesFromPPTX", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "StylesOpenStyleSheet", "Open slide master stylesheet", "OpenStyleSheet", "Styles", "Instrumenta > Styles > Inside splitbutton 'Open slide master stylesheet'", "Instrumenta [Text] > Styles > Inside splitbutton 'Open slide master stylesheet'"
    AddFeature "StylesCreateStyleSheetLayout", "Create stylesheet on master of current slide", "CreateStyleSheetLayout", "Styles", "Instrumenta > Styles > Inside splitbutton 'Open slide master stylesheet'", "Instrumenta [Text] > Styles > Inside splitbutton 'Open slide master stylesheet'"
    AddFeature "StylesCreateStyleSheetOnAllMasters", "Create stylesheets on all masters", "CreateStyleSheetOnAllMasters", "Styles", "Instrumenta > Styles > Inside splitbutton 'Open slide master stylesheet'", "Instrumenta [Text] > Styles > Inside splitbutton 'Open slide master stylesheet'"
    AddFeature "StylesRemoveStylesheet", "Remove stylesheet (current master)", "RemoveInstrumentaStylesheet", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "StylesRemoveStyleSheetsFromAllMasters", "Remove stylesheets (all masters)", "RemoveStyleSheetsFromAllMasters", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "StylesRemoveAllStyleTags", "Remove all style tags", "RemoveAllInstrumentaStyleTags", "Styles", "Instrumenta > Styles", "Instrumenta [Text] > Styles"
    AddFeature "CloneSelectionRight", "Clone selection to right", "ObjectsCloneRight", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "CloneSelectionDown", "Clone selection down", "ObjectsCloneDown", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "ObjectSetPositionDialog", "Set position", "InitialiseSetPositionAppEventHandler", "Shapes", "Instrumenta > Shapes > Inside splitbutton 'Set position'", "[Shapes] > Shapes > Inside splitbutton 'Set position'"
    AddFeature "ObjectsCopyRoundedCorner", "Copy rounded corner of first selected shape to selected shapes", "ObjectsCopyRoundedCorner", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "ObjectsCopyShapeTypeAndAdjustments", "Copy shape type and all adjustments of first selected shape to selected shapes", "ObjectsCopyShapeTypeAndAdjustments", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "RectifyLines", "Rectify lines", "RectifyLines", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "ConnectRectangleShapesRightToLeft", "Shape 1 right side to shape 2 left side", "ConnectRectangleShapesRightToLeft", "Shapes", "Instrumenta > Shapes > Inside menu 'Connect sides of 2 rectangles'", "[Shapes] > Shapes > Inside menu 'Connect sides of 2 rectangles'"
    AddFeature "ConnectRectangleShapesLeftToRight", "Shape 1 left side to shape 2 right side", "ConnectRectangleShapesLeftToRight", "Shapes", "Instrumenta > Shapes > Inside menu 'Connect sides of 2 rectangles'", "[Shapes] > Shapes > Inside menu 'Connect sides of 2 rectangles'"
    AddFeature "ConnectRectangleShapesBottomToTop", "Shape 1 bottom side to shape 2 top side", "ConnectRectangleShapesBottomToTop", "Shapes", "Instrumenta > Shapes > Inside menu 'Connect sides of 2 rectangles'", "[Shapes] > Shapes > Inside menu 'Connect sides of 2 rectangles'"
    AddFeature "ConnectRectangleShapesTopToBottom", "Shape 1 top side to shape 2 bottom side", "ConnectRectangleShapesTopToBottom", "Shapes", "Instrumenta > Shapes > Inside menu 'Connect sides of 2 rectangles'", "[Shapes] > Shapes > Inside menu 'Connect sides of 2 rectangles'"
    AddFeature "IncreaseShapeTransparency", "Increase shape transparency", "IncreaseShapeTransparency", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "LockAspectRatioToggleSelectedShapes", "Toggle lock aspect ratio of selected shape(s)", "LockAspectRatioToggleSelectedShapes", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "MoveSelectedShapeToMaster", "Move selected shape to master slide", "MoveSelectedShapeToMaster", "Shapes", "Instrumenta > Shapes > Inside menu 'Move shapes to master'", "[Shapes] > Shapes > Inside menu 'Move shapes to master'"
    AddFeature "CopySelectedShapeToMaster", "Copy selected shape to master slide", "CopySelectedShapeToMaster", "Shapes", "Instrumenta > Shapes > Inside menu 'Move shapes to master'", "[Shapes] > Shapes > Inside menu 'Move shapes to master'"
    AddFeature "MoveSelectedShapeToAllMasters", "Move selected shape to all master slides", "MoveSelectedShapeToAllMasters", "Shapes", "Instrumenta > Shapes > Inside menu 'Move shapes to master'", "[Shapes] > Shapes > Inside menu 'Move shapes to master'"
    AddFeature "CopySelectedShapeToAllMasters", "Copy selected shape to all master slides", "CopySelectedShapeToAllMasters", "Shapes", "Instrumenta > Shapes > Inside menu 'Move shapes to master'", "[Shapes] > Shapes > Inside menu 'Move shapes to master'"
    AddFeature "MoveSelectedShapeToUsedMasters", "Move selected shape to all master slides currently in use", "MoveSelectedShapeToUsedMasters", "Shapes", "Instrumenta > Shapes > Inside menu 'Move shapes to master'", "[Shapes] > Shapes > Inside menu 'Move shapes to master'"
    AddFeature "CopySelectedShapeToUsedMasters", "Copy selected shape to all master slides currently in use", "CopySelectedShapeToUsedMasters", "Shapes", "Instrumenta > Shapes > Inside menu 'Move shapes to master'", "[Shapes] > Shapes > Inside menu 'Move shapes to master'"
    AddFeature "LockToggleSelectedShapesButton", "Toggle lock or unlock position of selected shapes", "LockToggleSelectedShapes", "Shapes", "Instrumenta > Shapes > Inside splitbutton 'Toggle lock or unlock position of selected shapes'", "[Shapes] > Shapes > Inside splitbutton 'Toggle lock or unlock position of selected shapes'"
    AddFeature "LockToggleAllShapesOnAllSlidesButton", "Toggle lock or unlock position all shapes on all slides", "LockToggleAllShapesOnAllSlides", "Shapes", "Instrumenta > Shapes > Inside splitbutton 'Toggle lock or unlock position of selected shapes'", "[Shapes] > Shapes > Inside splitbutton 'Toggle lock or unlock position of selected shapes'"
    AddFeature "LockAllShapesOnAllSlidesButton", "Lock position of all shapes on all slides", "LockAllShapesOnAllSlides", "Shapes", "Instrumenta > Shapes > Inside splitbutton 'Toggle lock or unlock position of selected shapes'", "[Shapes] > Shapes > Inside splitbutton 'Toggle lock or unlock position of selected shapes'"
    AddFeature "UnLockAllShapesOnAllSlides", "Unlock position of all shapes on all slides", "UnLockAllShapesOnAllSlides", "Shapes", "Instrumenta > Shapes > Inside splitbutton 'Toggle lock or unlock position of selected shapes'", "[Shapes] > Shapes > Inside splitbutton 'Toggle lock or unlock position of selected shapes'"
    AddFeature "GroupShapesByColumnsButton", "Group shapes by columns", "GroupShapesByColumns", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "GroupShapesByRowsButton", "Group shapes by rows", "GroupShapesByRows", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "SelectShapesByFillColor", "Select shapes with same fill color", "ObjectsSelectBySameFillColor", "Shapes", "Instrumenta > Shapes > Inside menu 'Select shape by attributes'", "[Shapes] > Shapes > Inside menu 'Select shape by attributes'"
    AddFeature "SelectShapesByLineColor", "Select shapes with same line color", "ObjectsSelectBySameLineColor", "Shapes", "Instrumenta > Shapes > Inside menu 'Select shape by attributes'", "[Shapes] > Shapes > Inside menu 'Select shape by attributes'"
    AddFeature "SelectShapesByFillAndLineColor", "Select shapes with same fill and line color", "ObjectsSelectBySameFillAndLineColor", "Shapes", "Instrumenta > Shapes > Inside menu 'Select shape by attributes'", "[Shapes] > Shapes > Inside menu 'Select shape by attributes'"
    AddFeature "SelectShapesBySameWidthAndHeight", "Select shapes with same size", "ObjectsSelectBySameWidthAndHeight", "Shapes", "Instrumenta > Shapes > Inside menu 'Select shape by attributes'", "[Shapes] > Shapes > Inside menu 'Select shape by attributes'"
    AddFeature "SelectShapesBySameWidth", "Select shapes with same width", "ObjectsSelectBySameWidth", "Shapes", "Instrumenta > Shapes > Inside menu 'Select shape by attributes'", "[Shapes] > Shapes > Inside menu 'Select shape by attributes'"
    AddFeature "SelectShapesBySameHeight", "Select shapes with same height", "ObjectsSelectBySameHeight", "Shapes", "Instrumenta > Shapes > Inside menu 'Select shape by attributes'", "[Shapes] > Shapes > Inside menu 'Select shape by attributes'"
    AddFeature "SelectShapesBySameType", "Select shapes with same type", "ObjectsSelectBySameType", "Shapes", "Instrumenta > Shapes > Inside menu 'Select shape by attributes'", "[Shapes] > Shapes > Inside menu 'Select shape by attributes'"
    AddFeature "CopyPosition", "Copy position and dimensions", "CopyPosition", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "PastePosition", "Paste position", "PastePosition", "Shapes", "Instrumenta > Shapes > Inside splitbutton 'Paste position'", "[Shapes] > Shapes > Inside splitbutton 'Paste position'"
    AddFeature "PastePositionAndDimensions", "Paste position and dimensions", "PastePositionAndDimensions", "Shapes", "Instrumenta > Shapes > Inside splitbutton 'Paste position'", "[Shapes] > Shapes > Inside splitbutton 'Paste position'"
    AddFeature "CreateMultiSlideShape", "Copy shape to multiple slides (multislide shape)", "ShowFormCopyShapeToMultipleSlides", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "UpdateMultiSlideShape", "Update position and dimensions of selected multislide shape on all slides", "UpdateTaggedShapePositionAndDimensions", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "DeleteMultislideShape", "Delete selected multislide shape on all slides", "DeleteTaggedShapes", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "RemoveMargins", "Remove margins", "ObjectsMarginsToZero", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "IncreaseMargins", "Increase margins", "ObjectsMarginsIncrease", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "DecreaseMargins", "Decrease margins", "ObjectsMarginsDecrease", "Shapes", "Instrumenta > Shapes", "[Shapes] > Shapes"
    AddFeature "ApplySameCropToSelectedImages", "Apply same crop to selected pictures", "ApplySameCropToSelectedImages", "Pictures", "Instrumenta > Pictures > Inside splitbutton 'PictureCrop'", "[Shapes] > Pictures > Inside splitbutton 'PictureCrop'"
    AddFeature "PictureCropToSlide", "Crop picture or shape to slide", "PictureCropToSlide", "Pictures", "Instrumenta > Pictures > Inside splitbutton 'PictureCrop'", "[Shapes] > Pictures > Inside splitbutton 'PictureCrop'"
    AddFeature "PictureCropToPadding", "Trim picture padding by edge color", "CropSelectedImageByDominantEdgeColor", "Pictures", "Instrumenta > Pictures > Inside splitbutton 'PictureCrop'", "[Shapes] > Pictures > Inside splitbutton 'PictureCrop'"
    AddFeature "ObjectsAlignLeftsButton", "Align left", "ObjectsAlignLefts", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsAlignCentersButton", "Align center", "ObjectsAlignCenters", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsAlignRightsButton", "Align right", "ObjectsAlignRights", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsAlignBottomsButton", "Align bottom", "ObjectsAlignBottoms", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsAlignMiddlesButton", "Align middle", "ObjectsAlignMiddles", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsAlignTopsButton", "Align top", "ObjectsAlignTops", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsDistributeHorizontallyButton", "Distribute horizontally", "ObjectsDistributeHorizontally", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsDistributeVerticallyButton", "Distribute vertically", "ObjectsDistributeVertically", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ResizeAndSpaceEvenHorizontal", "Resize and space horizontally (equal spacing)", "ResizeAndSpaceEvenHorizontal", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Resize and space evenly'", "[Shapes] > Align, distribute and size > Inside menu 'Resize and space evenly'"
    AddFeature "ResizeAndSpaceEvenHorizontalPreserveFirst", "Resize and space horizontally (preserve first)", "ResizeAndSpaceEvenHorizontalPreserveFirst", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Resize and space evenly'", "[Shapes] > Align, distribute and size > Inside menu 'Resize and space evenly'"
    AddFeature "ResizeAndSpaceEvenHorizontalPreserveLast", "Resize and space horizontally (preserve last)", "ResizeAndSpaceEvenHorizontalPreserveLast", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Resize and space evenly'", "[Shapes] > Align, distribute and size > Inside menu 'Resize and space evenly'"
    AddFeature "ResizeAndSpaceEvenVertical", "Resize and space vertically (equal spacing)", "ResizeAndSpaceEvenVertical", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Resize and space evenly'", "[Shapes] > Align, distribute and size > Inside menu 'Resize and space evenly'"
    AddFeature "ResizeAndSpaceEvenVerticalPreserveFirst", "Resize and space vertically (preserve first)", "ResizeAndSpaceEvenVerticalPreserveFirst", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Resize and space evenly'", "[Shapes] > Align, distribute and size > Inside menu 'Resize and space evenly'"
    AddFeature "ResizeAndSpaceEvenVerticalPreserveLast", "Resize and space vertically (preserve last)", "ResizeAndSpaceEvenVerticalPreserveLast", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Resize and space evenly'", "[Shapes] > Align, distribute and size > Inside menu 'Resize and space evenly'"
    AddFeature "ObjectsSwapPosition", "Swap position of two shapes", "ObjectsSwapPosition", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside splitbutton 'Swap position of two shapes'", "[Shapes] > Align, distribute and size > Inside splitbutton 'Swap position of two shapes'"
    AddFeature "ObjectsSwapPositionCentered", "Swap position of two shapes (centered)", "ObjectsSwapPositionCentered", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Swap position of two shapes' > Inside splitbutton 'Swap position of two shapes'", "[Shapes] > Align, distribute and size > Inside menu 'Swap position of two shapes' > Inside splitbutton 'Swap position of two shapes'"
    AddFeature "ObjectsSameHeightButton", "Set same height", "ObjectsSameHeight", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsSameWidthButton", "Set same width", "ObjectsSameWidth", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsSameHeightAndWidthButton", "Set same height and width", "ObjectsSameHeightAndWidth", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsSizeToTallest", "Size to tallest", "ObjectsSizeToTallest", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsSizeShortest", "Size to shortest", "ObjectsSizeToShortest", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsSizeToWidest", "Size to widest", "ObjectsSizeToWidest", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsSizeNarrowest", "Size to narrowest", "ObjectsSizeToNarrowest", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsRemoveSpaceHorizontal", "Remove horizontal gap between shapes (direction left)", "ObjectsRemoveSpacingHorizontal", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside splitbutton 'Remove horizontal gap between shapes (direction left)'", "[Shapes] > Align, distribute and size > Inside splitbutton 'Remove horizontal gap between shapes (direction left)'"
    AddFeature "ObjectsRemoveSpaceHorizontalRight", "Remove horizontal gap between shapes (direction right)", "ObjectsRemoveSpacingHorizontalRight", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside splitbutton 'Remove horizontal gap between shapes (direction left)'", "[Shapes] > Align, distribute and size > Inside splitbutton 'Remove horizontal gap between shapes (direction left)'"
    AddFeature "ObjectsIncreaseSpacingHorizontal", "Increase horizontal gap between shapes", "ObjectsIncreaseSpacingHorizontal", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsDecreaseSpacingHorizontal", "Decrease horizontal gap between shapes", "ObjectsDecreaseSpacingHorizontal", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsRemoveSpaceVertical", "Remove vertical gap between shapes (direction up)", "ObjectsRemoveSpacingVertical", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside splitbutton 'Remove vertical gap between shapes (direction up)'", "[Shapes] > Align, distribute and size > Inside splitbutton 'Remove vertical gap between shapes (direction up)'"
    AddFeature "ObjectsRemoveSpaceVerticalDown", "Remove vertical gap between shapes (direction down)", "ObjectsRemoveSpacingVerticalDown", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside splitbutton 'Remove vertical gap between shapes (direction up)'", "[Shapes] > Align, distribute and size > Inside splitbutton 'Remove vertical gap between shapes (direction up)'"
    AddFeature "ObjectsIncreaseSpacingVertical", "Increase vertical gap between shapes", "ObjectsIncreaseSpacingVertical", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsDecreaseSpacingVertical", "Decrease vertical gap between shapes", "ObjectsDecreaseSpacingVertical", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsArrangeShapesButton", "Arrange shapes", "ArrangeShapes", "Align, distribute and size", "Instrumenta > Align, distribute and size", "[Shapes] > Align, distribute and size"
    AddFeature "ObjectsAlignToTable", "Center align shapes to table cells", "ObjectsAlignToTable", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside splitbutton 'Center align shapes to table cells'", "[Shapes] > Align, distribute and size > Inside splitbutton 'Center align shapes to table cells'"
    AddFeature "ObjectsAlignToTableColumn", "Center align shapes to columns of a table", "ObjectsAlignToTableColumn", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside splitbutton 'Center align shapes to table cells'", "[Shapes] > Align, distribute and size > Inside splitbutton 'Center align shapes to table cells'"
    AddFeature "ObjectsAlignToTableRow", "Center align shapes to rows of a table", "ObjectsAlignToTableRow", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside splitbutton 'Center align shapes to table cells'", "[Shapes] > Align, distribute and size > Inside splitbutton 'Center align shapes to table cells'"
    AddFeature "ObjectsStretchLeft", "Stretch shapes to left", "ObjectsStretchLeft", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Stretch shapes'", "[Shapes] > Align, distribute and size > Inside menu 'Stretch shapes'"
    AddFeature "ObjectsStretchRight", "Stretch shapes to right", "ObjectsStretchRight", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Stretch shapes'", "[Shapes] > Align, distribute and size > Inside menu 'Stretch shapes'"
    AddFeature "ObjectsStretchTop", "Stretch shapes to top", "ObjectsStretchTop", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Stretch shapes'", "[Shapes] > Align, distribute and size > Inside menu 'Stretch shapes'"
    AddFeature "ObjectsStretchBottom", "Stretch shapes to bottom", "ObjectsStretchBottom", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Stretch shapes'", "[Shapes] > Align, distribute and size > Inside menu 'Stretch shapes'"
    AddFeature "ObjectsStretchLeftShapeRight", "Stretch shapes to the right edge of the leftmost shape", "ObjectsStretchLeftShapeRight", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Stretch shapes'", "[Shapes] > Align, distribute and size > Inside menu 'Stretch shapes'"
    AddFeature "ObjectsStretchRightShapeLeft", "Stretch shapes to the left edge of the rightmost shape", "ObjectsStretchRightShapeLeft", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Stretch shapes'", "[Shapes] > Align, distribute and size > Inside menu 'Stretch shapes'"
    AddFeature "ObjectsStretchTopShapeBottom", "Stretch shapes to the bottom edge of the topmost shape", "ObjectsStretchTopShapeBottom", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Stretch shapes'", "[Shapes] > Align, distribute and size > Inside menu 'Stretch shapes'"
    AddFeature "ObjectsStretchBottomShapeTop", "Stretch shapes to the top edge of the bottommost shape", "ObjectsStretchBottomShapeTop", "Align, distribute and size", "Instrumenta > Align, distribute and size > Inside menu 'Stretch shapes'", "[Shapes] > Align, distribute and size > Inside menu 'Stretch shapes'"
    AddFeature "InsertColumnToLeftKeepOtherColumnWidths", "Insert column left (preserve widths)", "InsertColumnToLeftKeepOtherColumnWidths", "Tables", "Instrumenta > Tables > Inside splitbutton 'TableInsertColumnsLeft'", "[Tables] > Tables > Inside splitbutton 'TableInsertColumnsLeft'"
    AddFeature "InsertColumnToRightKeepOtherColumnWidths", "Insert column right (preserve widths)", "InsertColumnToRightKeepOtherColumnWidths", "Tables", "Instrumenta > Tables > Inside splitbutton 'TableInsertColumnsRight'", "[Tables] > Tables > Inside splitbutton 'TableInsertColumnsRight'"
    AddFeature "TableDistributeRowsWithGaps", "Distribute rows ignoring row gaps", "TableDistributeRowsWithGaps", "Tables", "Instrumenta > Tables > Inside splitbutton 'TableRowsDistribute'", "[Tables] > Tables > Inside splitbutton 'TableRowsDistribute'"
    AddFeature "TableDistributeColumnsWithGaps", "Distribute columns ignoring column gaps", "TableDistributeColumnsWithGaps", "Tables", "Instrumenta > Tables > Inside splitbutton 'TableColumnsDistribute'", "[Tables] > Tables > Inside splitbutton 'TableColumnsDistribute'"
    AddFeature "TableQuickFormat", "Quick format table", "TableQuickFormat", "Tables", "Instrumenta > Tables > Inside splitbutton 'Quick format table'", "[Tables] > Tables > Inside splitbutton 'Quick format table'"
    AddFeature "TableRemoveBackgrounds", "Remove cell fills", "TableRemoveBackgrounds", "Tables", "Instrumenta > Tables > Inside splitbutton 'Quick format table'", "[Tables] > Tables > Inside splitbutton 'Quick format table'"
    AddFeature "TableRemoveBorders", "Remove all borders", "TableRemoveBorders", "Tables", "Instrumenta > Tables > Inside splitbutton 'Quick format table'", "[Tables] > Tables > Inside splitbutton 'Quick format table'"
    AddFeature "TableConvertTableToShapes", "Convert table to shapes", "ConvertTableToShapes", "Tables", "Instrumenta > Tables", "[Tables] > Tables"
    AddFeature "TableConvertShapesToTable", "Convert shapes to table", "ConvertShapesToTable", "Tables", "Instrumenta > Tables", "[Tables] > Tables"
    AddFeature "TableTranspose", "Transpose table", "TableTranspose", "Tables", "Instrumenta > Tables", "[Tables] > Tables"
    AddFeature "SplitTableByRowButton", "Split table by row", "SplitTableByRow", "Tables", "Instrumenta > Tables > Inside splitbutton 'Split table by row'", "[Tables] > Tables > Inside splitbutton 'Split table by row'"
    AddFeature "SplitTableByColumnButton", "Split table by column", "SplitTableByColumn", "Tables", "Instrumenta > Tables > Inside splitbutton 'Split table by row'", "[Tables] > Tables > Inside splitbutton 'Split table by row'"
    AddFeature "TableSumButton", "Sum column (values above selected cells)", "TableSum", "Tables", "Instrumenta > Tables > Inside splitbutton 'Sum column (values above selected cells)'", "[Tables] > Tables > Inside splitbutton 'Sum column (values above selected cells)'"
    AddFeature "TableRowSumButton", "Sum row (values left from selected cells)", "TableRowSum", "Tables", "Instrumenta > Tables > Inside splitbutton 'Sum column (values above selected cells)'", "[Tables] > Tables > Inside splitbutton 'Sum column (values above selected cells)'"
    AddFeature "MoveTableRowUp", "Move row up", "MoveTableRowUp", "Tables", "Instrumenta > Tables > Inside menu 'Move rows and columns'", "[Tables] > Tables > Inside menu 'Move rows and columns'"
    AddFeature "MoveTableRowDown", "Move row down", "MoveTableRowDown", "Tables", "Instrumenta > Tables > Inside menu 'Move rows and columns'", "[Tables] > Tables > Inside menu 'Move rows and columns'"
    AddFeature "MoveTableColumnLeft", "Move column left", "MoveTableColumnLeft", "Tables", "Instrumenta > Tables > Inside menu 'Move rows and columns'", "[Tables] > Tables > Inside menu 'Move rows and columns'"
    AddFeature "MoveTableColumnRight", "Move column right", "MoveTableColumnRight", "Tables", "Instrumenta > Tables > Inside menu 'Move rows and columns'", "[Tables] > Tables > Inside menu 'Move rows and columns'"
    AddFeature "MoveTableRowUpIgnoreBorders", "Move row up (ignore borders)", "MoveTableRowUpIgnoreBorders", "Tables", "Instrumenta > Tables > Inside menu 'Move rows and columns'", "[Tables] > Tables > Inside menu 'Move rows and columns'"
    AddFeature "MoveTableRowDownIgnoreBorders", "Move row down (ignore borders)", "MoveTableRowDownIgnoreBorders", "Tables", "Instrumenta > Tables > Inside menu 'Move rows and columns'", "[Tables] > Tables > Inside menu 'Move rows and columns'"
    AddFeature "MoveTableColumnLeftIgnoreBorders", "Move column left (ignore borders)", "MoveTableColumnLeftIgnoreBorders", "Tables", "Instrumenta > Tables > Inside menu 'Move rows and columns'", "[Tables] > Tables > Inside menu 'Move rows and columns'"
    AddFeature "MoveTableColumnRightIgnoreBorders", "Move column right (ignore borders)", "MoveTableColumnRightIgnoreBorders", "Tables", "Instrumenta > Tables > Inside menu 'Move rows and columns'", "[Tables] > Tables > Inside menu 'Move rows and columns'"
    AddFeature "MoveTableRowUpTextOnly", "Move row up (text only)", "MoveTableRowUpTextOnly", "Tables", "Instrumenta > Tables > Inside menu 'Move rows and columns'", "[Tables] > Tables > Inside menu 'Move rows and columns'"
    AddFeature "MoveTableRowDownTextOnly", "Move row down (text only)", "MoveTableRowDownTextOnly", "Tables", "Instrumenta > Tables > Inside menu 'Move rows and columns'", "[Tables] > Tables > Inside menu 'Move rows and columns'"
    AddFeature "MoveTableColumnLeftTextOnly", "Move column left (text only)", "MoveTableColumnLeftTextOnly", "Tables", "Instrumenta > Tables > Inside menu 'Move rows and columns'", "[Tables] > Tables > Inside menu 'Move rows and columns'"
    AddFeature "MoveTableColumnRightTextOnly", "Move column right (text only)", "MoveTableColumnRightTextOnly", "Tables", "Instrumenta > Tables > Inside menu 'Move rows and columns'", "[Tables] > Tables > Inside menu 'Move rows and columns'"
    AddFeature "OptimizeTableHeightQuick", "Quick", "OptimizeTableHeightQuick", "Tables", "Instrumenta > Tables > Inside menu 'Optimize table height while preserving width'", "[Tables] > Tables > Inside menu 'Optimize table height while preserving width'"
    AddFeature "OptimizeTableHeight3Iterations", "Optimized (3 iterations)", "OptimizeTableHeight3Iterations", "Tables", "Instrumenta > Tables > Inside menu 'Optimize table height while preserving width'", "[Tables] > Tables > Inside menu 'Optimize table height while preserving width'"
    AddFeature "OptimizeTableHeight5Iterations", "Optimized (5 iterations)", "OptimizeTableHeight5Iterations", "Tables", "Instrumenta > Tables > Inside menu 'Optimize table height while preserving width'", "[Tables] > Tables > Inside menu 'Optimize table height while preserving width'"
    AddFeature "OptimizeTableHeight10Iterations", "Optimized (10 iterations)", "OptimizeTableHeight10Iterations", "Tables", "Instrumenta > Tables > Inside menu 'Optimize table height while preserving width'", "[Tables] > Tables > Inside menu 'Optimize table height while preserving width'"
    AddFeature "OptimizeTableHeight20Iterations", "Optimized (20 iterations)", "OptimizeTableHeight20Iterations", "Tables", "Instrumenta > Tables > Inside menu 'Optimize table height while preserving width'", "[Tables] > Tables > Inside menu 'Optimize table height while preserving width'"
    AddFeature "TableColumnGapsEven", "Add column gaps", "TableColumnGapsEven", "Tables", "Instrumenta > Tables > Inside splitbutton 'Add column gaps'", "[Tables] > Tables > Inside splitbutton 'Add column gaps'"
    AddFeature "TableColumnGapsOdd", "Add column gaps (including left and right sides)", "TableColumnGapsOdd", "Tables", "Instrumenta > Tables > Inside splitbutton 'Add column gaps'", "[Tables] > Tables > Inside splitbutton 'Add column gaps'"
    AddFeature "TableColumnRemoveGaps", "Remove column gaps", "TableColumnRemoveGaps", "Tables", "Instrumenta > Tables > Inside splitbutton 'Add column gaps'", "[Tables] > Tables > Inside splitbutton 'Add column gaps'"
    AddFeature "TableColumnIncreaseGaps", "Increase column gaps", "TableColumnIncreaseGaps", "Tables", "Instrumenta > Tables", "[Tables] > Tables"
    AddFeature "TableColumnDecreaseGaps", "Decrease column gaps", "TableColumnDecreaseGaps", "Tables", "Instrumenta > Tables", "[Tables] > Tables"
    AddFeature "TableRowGapsEven", "Add row gaps", "TableRowGapsEven", "Tables", "Instrumenta > Tables > Inside splitbutton 'Add row gaps'", "[Tables] > Tables > Inside splitbutton 'Add row gaps'"
    AddFeature "TableRowGapsOdd", "Add row gaps (including top and bottom sides)", "TableRowGapsOdd", "Tables", "Instrumenta > Tables > Inside splitbutton 'Add row gaps'", "[Tables] > Tables > Inside splitbutton 'Add row gaps'"
    AddFeature "TableRowRemoveGaps", "Remove row gaps", "TableRowRemoveGaps", "Tables", "Instrumenta > Tables > Inside splitbutton 'Add row gaps'", "[Tables] > Tables > Inside splitbutton 'Add row gaps'"
    AddFeature "TableRowIncreaseGaps", "Increase row gaps", "TableRowIncreaseGaps", "Tables", "Instrumenta > Tables", "[Tables] > Tables"
    AddFeature "TableRowDecreaseGaps", "Decrease row gaps", "TableRowDecreaseGaps", "Tables", "Instrumenta > Tables", "[Tables] > Tables"
    AddFeature "TableRemoveMargins", "Remove margins of selected table or selected cells", "TablesMarginsToZero", "Tables", "Instrumenta > Tables", "[Tables] > Tables"
    AddFeature "TableIncreaseMargins", "Increase margins of selected table or selected cells", "TablesMarginsIncrease", "Tables", "Instrumenta > Tables", "[Tables] > Tables"
    AddFeature "TableDecreaseMargins", "Decrease margins of selected table or selected cells", "TablesMarginsDecrease", "Tables", "Instrumenta > Tables", "[Tables] > Tables"
    AddFeature "SaveSelectedSlidesButton", "Save selected slides", "SaveSelectedSlides", "Export", "Instrumenta > Export", "[Advanced] > Export"
    AddFeature "EmailSelectedSlidesButton", "E-mail selected slides", "EmailSelectedSlides", "Export", "Instrumenta > Export", "[Advanced] > Export"
    AddFeature "EmailSelectedSlidesAsPDFButton", "E-mail selected slides as PDF", "EmailSelectedSlidesAsPDF", "Export", "Instrumenta > Export", "[Advanced] > Export"
    AddFeature "CopySlideNotesToWord", "Export slide notes to Word", "CopySlideNotesToWord", "Export", "Instrumenta > Export > Inside menu 'Copy storyline/slide note'", "[Advanced] > Export > Inside menu 'Copy storyline/slide note'"
    AddFeature "CopySlideNotesToClipboard", "Copy slide notes to clipboard", "CopySlideNotesToClipboardOnly", "Export", "Instrumenta > Export > Inside menu 'Copy storyline/slide note'", "[Advanced] > Export > Inside menu 'Copy storyline/slide note'"
    AddFeature "CopyStorylineToWord", "Export storyline to Word", "CopyStorylineToWord", "Export", "Instrumenta > Export > Inside menu 'Copy storyline/slide note'", "[Advanced] > Export > Inside menu 'Copy storyline/slide note'"
    AddFeature "CopyStorylineToClipboard", "Copy storyline to clipboard", "CopyStorylineToClipBoardOnly", "Export", "Instrumenta > Export > Inside menu 'Copy storyline/slide note'", "[Advanced] > Export > Inside menu 'Copy storyline/slide note'"
    AddFeature "PasteStorylineInSelectedShapeButton", "Paste storyline in shape", "PasteStorylineInSelectedShape", "Export", "Instrumenta > Export > Inside menu 'Copy storyline/slide note'", "[Advanced] > Export > Inside menu 'Copy storyline/slide note'"
    AddFeature "InsertSlideFromSlideLibraryButton", "Insert slide from slide library", "ShowSlideLibrary", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Insert slide from slide library'", "[Advanced] > Paste and insert > Inside splitbutton 'Insert slide from slide library'"
    AddFeature "OpenSlideLibraryFileButton", "Open slide library file", "OpenSlideLibraryFile", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Insert slide from slide library'", "[Advanced] > Paste and insert > Inside splitbutton 'Insert slide from slide library'"
    AddFeature "AddSelectedSlidesToLibraryFileButton", "Copy selected slides to slide library", "AddSelectedSlidesToLibraryFile", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Insert slide from slide library'", "[Advanced] > Paste and insert > Inside splitbutton 'Insert slide from slide library'"
    AddFeature "AgendaPages", "Create or update agenda pages", "CreateOrUpdateMasterAgenda", "Paste and insert", "Instrumenta > Paste and insert", "[Advanced] > Paste and insert"
    AddFeature "EmojiGallery1", "Smileys", "EmojiGallery_OnAction", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Emoji'", "[Advanced] > Paste and insert > Inside menu 'Emoji'"
    AddFeature "EmojiGallery2", "Gestures and body parts", "EmojiGallery_OnAction", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Emoji'", "[Advanced] > Paste and insert > Inside menu 'Emoji'"
    AddFeature "EmojiGallery3", "People and fantasy", "EmojiGallery_OnAction", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Emoji'", "[Advanced] > Paste and insert > Inside menu 'Emoji'"
    AddFeature "EmojiGallery4", "Clothing and accessories", "EmojiGallery_OnAction", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Emoji'", "[Advanced] > Paste and insert > Inside menu 'Emoji'"
    AddFeature "EmojiGallery5", "Animals and nature", "EmojiGallery_OnAction", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Emoji'", "[Advanced] > Paste and insert > Inside menu 'Emoji'"
    AddFeature "EmojiGallery6", "Food and drink", "EmojiGallery_OnAction", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Emoji'", "[Advanced] > Paste and insert > Inside menu 'Emoji'"
    AddFeature "EmojiGallery7", "Activity and sports", "EmojiGallery_OnAction", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Emoji'", "[Advanced] > Paste and insert > Inside menu 'Emoji'"
    AddFeature "EmojiGallery8", "Travel and places", "EmojiGallery_OnAction", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Emoji'", "[Advanced] > Paste and insert > Inside menu 'Emoji'"
    AddFeature "EmojiGallery9", "Objects", "EmojiGallery_OnAction", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Emoji'", "[Advanced] > Paste and insert > Inside menu 'Emoji'"
    AddFeature "EmojiGallery10", "Symbols", "EmojiGallery_OnAction", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Emoji'", "[Advanced] > Paste and insert > Inside menu 'Emoji'"
    AddFeature "HarveyBall10", "10%", "GenerateHarveyBall10", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "HarveyBall20", "20%", "GenerateHarveyBall20", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "HarveyBall25", "25%", "GenerateHarveyBall25", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "HarveyBall30", "30%", "GenerateHarveyBall30", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "HarveyBall33", "33%", "GenerateHarveyBall33", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "HarveyBall40", "40%", "GenerateHarveyBall40", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "HarveyBall50", "50%", "GenerateHarveyBall50", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "HarveyBall60", "60%", "GenerateHarveyBall60", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "HarveyBall67", "67%", "GenerateHarveyBall67", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "HarveyBall70", "70%", "GenerateHarveyBall70", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "HarveyBall75", "75%", "GenerateHarveyBall75", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "HarveyBall80", "80%", "GenerateHarveyBall80", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "HarveyBall90", "90%", "GenerateHarveyBall90", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "HarveyBall100", "100%", "GenerateHarveyBall100", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "HarveyBallCustom", "Custom...", "GenerateHarveyBallCustom", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "AverageHarveyBall", "Average based on selected Harvey Balls", "AverageHarveyBall", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Harvey Balls'", "[Advanced] > Paste and insert > Inside menu 'Harvey Balls'"
    AddFeature "FivePointStar05", "0.5 star", "GenerateFivePointStars05", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Star rating'", "[Advanced] > Paste and insert > Inside menu 'Star rating'"
    AddFeature "FivePointStar10", "1 star", "GenerateFivePointStars10", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Star rating'", "[Advanced] > Paste and insert > Inside menu 'Star rating'"
    AddFeature "FivePointStar15", "1.5 star", "GenerateFivePointStars15", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Star rating'", "[Advanced] > Paste and insert > Inside menu 'Star rating'"
    AddFeature "FivePointStar20", "2 star", "GenerateFivePointStars20", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Star rating'", "[Advanced] > Paste and insert > Inside menu 'Star rating'"
    AddFeature "FivePointStar25", "2.5 star", "GenerateFivePointStars25", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Star rating'", "[Advanced] > Paste and insert > Inside menu 'Star rating'"
    AddFeature "FivePointStar30", "3 star", "GenerateFivePointStars30", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Star rating'", "[Advanced] > Paste and insert > Inside menu 'Star rating'"
    AddFeature "FivePointStar35", "3.5 star", "GenerateFivePointStars35", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Star rating'", "[Advanced] > Paste and insert > Inside menu 'Star rating'"
    AddFeature "FivePointStar40", "4 star", "GenerateFivePointStars40", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Star rating'", "[Advanced] > Paste and insert > Inside menu 'Star rating'"
    AddFeature "FivePointStar45", "4.5 star", "GenerateFivePointStars45", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Star rating'", "[Advanced] > Paste and insert > Inside menu 'Star rating'"
    AddFeature "FivePointStar50", "5 star", "GenerateFivePointStars50", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Star rating'", "[Advanced] > Paste and insert > Inside menu 'Star rating'"
    AddFeature "AverageFivePointStars", "Average based on selected star ratings", "AverageFivePointStars", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Star rating'", "[Advanced] > Paste and insert > Inside menu 'Star rating'"
    AddFeature "RAGStatusRed", "Red", "GenerateRAGStatusRed", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'RAG status'", "[Advanced] > Paste and insert > Inside menu 'RAG status'"
    AddFeature "RAGStatusAmber", "Amber", "GenerateRAGStatusAmber", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'RAG status'", "[Advanced] > Paste and insert > Inside menu 'RAG status'"
    AddFeature "RAGStatusGreen", "Green", "GenerateRAGStatusGreen", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'RAG status'", "[Advanced] > Paste and insert > Inside menu 'RAG status'"
    AddFeature "AverageRAGStatus", "Average based on selected RAG status-shapes", "AverageRAGStatus", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'RAG status'", "[Advanced] > Paste and insert > Inside menu 'RAG status'"
    AddFeature "LegendSquareVerticalThree", "3 squares vertical", "InsertLegendSquareVerticalThree", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Legend'", "[Advanced] > Paste and insert > Inside menu 'Legend'"
    AddFeature "LegendSquareVerticalFive", "5 squares vertical", "InsertLegendSquareVerticalFive", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Legend'", "[Advanced] > Paste and insert > Inside menu 'Legend'"
    AddFeature "LegendSquareVerticalTen", "10 squares vertical", "InsertLegendSquareVerticalTen", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Legend'", "[Advanced] > Paste and insert > Inside menu 'Legend'"
    AddFeature "LegendCircleVerticalThreed", "3 circles vertical", "InsertLegendCircleVerticalThree", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Legend'", "[Advanced] > Paste and insert > Inside menu 'Legend'"
    AddFeature "LegendCircleVerticalFive", "5 circles vertical", "InsertLegendCircleVerticalFive", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Legend'", "[Advanced] > Paste and insert > Inside menu 'Legend'"
    AddFeature "LegendCircleVerticalTen", "10 circles vertical", "InsertLegendCircleVerticalTen", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Legend'", "[Advanced] > Paste and insert > Inside menu 'Legend'"
    AddFeature "LegendSquareHorizontalThree", "3 squares horizontal", "InsertLegendSquareHorizontalThree", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Legend'", "[Advanced] > Paste and insert > Inside menu 'Legend'"
    AddFeature "LegendSquareHorizontalFive", "5 squares horizontal", "InsertLegendSquareHorizontalFive", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Legend'", "[Advanced] > Paste and insert > Inside menu 'Legend'"
    AddFeature "LegendSquareHorizontalTen", "10 squares horizontal", "InsertLegendSquareHorizontalTen", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Legend'", "[Advanced] > Paste and insert > Inside menu 'Legend'"
    AddFeature "LegendCircleHorizontalThree", "3 circles horizontal", "InsertLegendCircleHorizontalThree", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Legend'", "[Advanced] > Paste and insert > Inside menu 'Legend'"
    AddFeature "LegendCircleHorizontalFive", "5 circles horizontal", "InsertLegendCircleHorizontalFive", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Legend'", "[Advanced] > Paste and insert > Inside menu 'Legend'"
    AddFeature "LegendCircleHorizontalTen", "10 circles horizontal", "InsertLegendCircleHorizontalTen", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Legend'", "[Advanced] > Paste and insert > Inside menu 'Legend'"
    AddFeature "LegendInsertCustom", "Custom", "ShowCustomInsertLegend", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Legend'", "[Advanced] > Paste and insert > Inside menu 'Legend'"
    AddFeature "NewCaptionButton", "Insert caption for selected shape", "InsertCaption", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Insert caption for selected shape'", "[Advanced] > Paste and insert > Inside splitbutton 'Insert caption for selected shape'"
    AddFeature "ReNumberCaptionButton", "Renumber all captions", "ReNumberCaptions", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Insert caption for selected shape'", "[Advanced] > Paste and insert > Inside splitbutton 'Insert caption for selected shape'"
    AddFeature "NewNoteButton", "Note", "GenerateStickyNote", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Note'", "[Advanced] > Paste and insert > Inside splitbutton 'Note'"
    AddFeature "NotesMoveOffSlide", "Move notes off this slide", "MoveStickyNotesOffSlide", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Note'", "[Advanced] > Paste and insert > Inside splitbutton 'Note'"
    AddFeature "NotesMoveOnSlide", "Move notes on this slide", "MoveStickyNotesOnSlide", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Note'", "[Advanced] > Paste and insert > Inside splitbutton 'Note'"
    AddFeature "DeleteNotesOnSlide", "Delete notes on this slide", "DeleteStickyNotesOnSlide", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Note'", "[Advanced] > Paste and insert > Inside splitbutton 'Note'"
    AddFeature "ConvertCommentsToStickyNotes", "Convert comments on this slide to notes", "ConvertCommentsToStickyNotes", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Note'", "[Advanced] > Paste and insert > Inside splitbutton 'Note'"
    AddFeature "NotesMoveOffAllSlides", "Move notes off all slides", "MoveStickyNotesOffAllSlides", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Note'", "[Advanced] > Paste and insert > Inside splitbutton 'Note'"
    AddFeature "NotesMoveOnAllSlides", "Move notes on all slides", "MoveStickyNotesOnAllSlides", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Note'", "[Advanced] > Paste and insert > Inside splitbutton 'Note'"
    AddFeature "DeleteNotesOnAllSlides", "Delete notes on all slides", "DeleteStickyNotesOnAllSlides", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Note'", "[Advanced] > Paste and insert > Inside splitbutton 'Note'"
    AddFeature "ConvertAllCommentsToStickyNotes", "Convert comments on all slides to notes", "ConvertAllCommentsToStickyNotes", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Note'", "[Advanced] > Paste and insert > Inside splitbutton 'Note'"
    AddFeature "NewStepsCounterButton", "Steps counter", "GenerateStepsCounter", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Steps counter'", "[Advanced] > Paste and insert > Inside splitbutton 'Steps counter'"
    AddFeature "NewCrossSlideStepsCounterButton", "Add cross-slide steps counter", "GenerateCrossSlideStepsCounter", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Steps counter'", "[Advanced] > Paste and insert > Inside splitbutton 'Steps counter'"
    AddFeature "SelectAllStepsCounter", "Select all step counters on this slide", "SelectAllStepsCounter", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Steps counter'", "[Advanced] > Paste and insert > Inside splitbutton 'Steps counter'"
    AddFeature "SelectAllCrossSlideStepsCounter", "Select all cross-slide step counters on this slide", "SelectAllCrossSlideStepsCounter", "Paste and insert", "Instrumenta > Paste and insert > Inside splitbutton 'Steps counter'", "[Advanced] > Paste and insert > Inside splitbutton 'Steps counter'"
    AddFeature "StampConfidential", "CONFIDENTIAL", "GenerateStampConfidential", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Stamps'", "[Advanced] > Paste and insert > Inside menu 'Stamps'"
    AddFeature "StampDoNotDistribute", "DO NOT DISTRIBUTE", "GenerateStampDoNotDistribute", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Stamps'", "[Advanced] > Paste and insert > Inside menu 'Stamps'"
    AddFeature "StampDraft", "DRAFT", "GenerateStampDraft", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Stamps'", "[Advanced] > Paste and insert > Inside menu 'Stamps'"
    AddFeature "StampUpdated", "UPDATED", "GenerateStampUpdated", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Stamps'", "[Advanced] > Paste and insert > Inside menu 'Stamps'"
    AddFeature "StampNew", "NEW", "GenerateStampNew", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Stamps'", "[Advanced] > Paste and insert > Inside menu 'Stamps'"
    AddFeature "StampToBeRemoved", "TO BE REMOVED", "GenerateStampToBeRemoved", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Stamps'", "[Advanced] > Paste and insert > Inside menu 'Stamps'"
    AddFeature "StampToAppendix", "TO APPENDIX", "GenerateStampToAppendix", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Stamps'", "[Advanced] > Paste and insert > Inside menu 'Stamps'"
    AddFeature "StampsMoveOffSlide", "Move Stamps off this slide", "MoveStampsOffSlide", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Stamps'", "[Advanced] > Paste and insert > Inside menu 'Stamps'"
    AddFeature "StampsMoveOnSlide", "Move Stamps on this slide", "MoveStampsOnSlide", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Stamps'", "[Advanced] > Paste and insert > Inside menu 'Stamps'"
    AddFeature "DeleteStampsOnSlide", "Delete Stamps on this slide", "DeleteStampsOnSlide", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Stamps'", "[Advanced] > Paste and insert > Inside menu 'Stamps'"
    AddFeature "StampsMoveOffAllSlides", "Move Stamps off all slides", "MoveStampsOffAllSlides", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Stamps'", "[Advanced] > Paste and insert > Inside menu 'Stamps'"
    AddFeature "StampsMoveOnAllSlides", "Move Stamps on all slides", "MoveStampsOnAllSlides", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Stamps'", "[Advanced] > Paste and insert > Inside menu 'Stamps'"
    AddFeature "DeleteStampsOnAllSlides", "Delete Stamps on all slides", "DeleteStampsOnAllSlides", "Paste and insert", "Instrumenta > Paste and insert > Inside menu 'Stamps'", "[Advanced] > Paste and insert > Inside menu 'Stamps'"
    AddFeature "InsertProcessSmartArt", "Insert process (SmartArt)", "InsertProcessSmartArt", "Paste and insert", "Instrumenta > Paste and insert", "[Advanced] > Paste and insert"
    AddFeature "InsertQRCodeButton", "Insert QR-code", "InsertQRCode", "Paste and insert", "Instrumenta > Paste and insert", "[Advanced] > Paste and insert"
    AddFeature "ShowInstrumentaScriptButton", "Instrumenta script editor", "ShowScriptEditor", "Instrumenta script", "Instrumenta > Instrumenta script", "[Advanced] > Instrumenta script"
    AddFeature "InstrumentaScriptPresets", "Preset", "ScriptPreset_OnAction", "Instrumenta script", "Instrumenta > Instrumenta script", "[Advanced] > Instrumenta script"
    AddFeature "InstrumentaScriptPresetRun", "Run preset", "ScriptPreset_Run", "Instrumenta script", "Instrumenta > Instrumenta script", "[Advanced] > Instrumenta script"
    AddFeature "cleaning0", "Move to end and hide selected slides", "CleanUpHideAndMoveSelectedSlides", "Advanced", "Instrumenta > Advanced > Inside menu 'Cleaning tools'", "[Advanced] > Advanced > Inside menu 'Cleaning tools'"
    AddFeature "cleaning10", "Add page numbers to all/selected slides (except the first)", "CleanUpAddSlideNumbers", "Advanced", "Instrumenta > Advanced > Inside menu 'Cleaning tools'", "[Advanced] > Advanced > Inside menu 'Cleaning tools'"
    AddFeature "cleaning1", "Remove animations from all/selected slides", "CleanUpRemoveAnimationsFromAllSlides", "Advanced", "Instrumenta > Advanced > Inside menu 'Cleaning tools'", "[Advanced] > Advanced > Inside menu 'Cleaning tools'"
    AddFeature "cleaning2", "Remove entry transitions from all/selected slides", "CleanUpRemoveSlideShowTransitionsFromAllSlides", "Advanced", "Instrumenta > Advanced > Inside menu 'Cleaning tools'", "[Advanced] > Advanced > Inside menu 'Cleaning tools'"
    AddFeature "cleaning3", "Remove speaker notes from all/selected slides", "CleanUpRemoveSpeakerNotesFromAllSlides", "Advanced", "Instrumenta > Advanced > Inside menu 'Cleaning tools'", "[Advanced] > Advanced > Inside menu 'Cleaning tools'"
    AddFeature "cleaning4", "Remove comments from all/selected slides", "CleanUpRemoveCommentsFromAllSlides", "Advanced", "Instrumenta > Advanced > Inside menu 'Cleaning tools'", "[Advanced] > Advanced > Inside menu 'Cleaning tools'"
    AddFeature "cleaning5", "Remove all unused master slides", "CleanUpRemoveUnusedMasterSlides", "Advanced", "Instrumenta > Advanced > Inside menu 'Cleaning tools'", "[Advanced] > Advanced > Inside menu 'Cleaning tools'"
    AddFeature "cleaning6", "Remove all hidden slides", "CleanUpRemoveHiddenSlides", "Advanced", "Instrumenta > Advanced > Inside menu 'Cleaning tools'", "[Advanced] > Advanced > Inside menu 'Cleaning tools'"
    AddFeature "cleaning7", "Convert all/selected slides to pictures (readonly)", "ConvertSlidesToPictures", "Advanced", "Instrumenta > Advanced > Inside menu 'Cleaning tools'", "[Advanced] > Advanced > Inside menu 'Cleaning tools'"
    AddFeature "cleaning8", "Watermark and then convert all/selected slides to pictures (readonly)", "InsertWatermarkAndConvertSlidesToPictures", "Advanced", "Instrumenta > Advanced > Inside menu 'Cleaning tools'", "[Advanced] > Advanced > Inside menu 'Cleaning tools'"
    AddFeature "cleaning9", "Anonymize all/selected slides with Lorem Ipsum", "AnonymizeWithLoremIpsum", "Advanced", "Instrumenta > Advanced > Inside menu 'Cleaning tools'", "[Advanced] > Advanced > Inside menu 'Cleaning tools'"
    AddFeature "ReplaceColorsShapesButton", "Replace colors in selected shapes", "ScanColorsInSelectedShapes", "Advanced", "Instrumenta > Advanced > Inside splitbutton 'Replace colors in selected shapes'", "[Advanced] > Advanced > Inside splitbutton 'Replace colors in selected shapes'"
    AddFeature "ReplaceColorsButton", "Replace colors in all/selected slides", "ScanAndManageColors", "Advanced", "Instrumenta > Advanced > Inside splitbutton 'Replace colors in selected shapes'", "[Advanced] > Advanced > Inside splitbutton 'Replace colors in selected shapes'"
    AddFeature "PyramidBuilder", "Pyramid storyline builder", "ShowPyramidBuilder", "Advanced", "Instrumenta > Advanced", "[Advanced] > Advanced"
    AddFeature "ShowFormManageTagsButton", "Manage hidden tags to shapes or slides", "ShowFormManageTags", "Advanced", "Instrumenta > Advanced > Inside splitbutton 'Manage hidden tags to shapes or slides'", "[Advanced] > Advanced > Inside splitbutton 'Manage hidden tags to shapes or slides'"
    AddFeature "ShowTagsOnSlide", "Show slidetags on all slides", "ShowTagsOnSlide", "Advanced", "Instrumenta > Advanced > Inside splitbutton 'Manage hidden tags to shapes or slides'", "[Advanced] > Advanced > Inside splitbutton 'Manage hidden tags to shapes or slides'"
    AddFeature "HideTagsOnSlide", "Hide slidetags on all slides", "HideTagsOnSlide", "Advanced", "Instrumenta > Advanced > Inside splitbutton 'Manage hidden tags to shapes or slides'", "[Advanced] > Advanced > Inside splitbutton 'Manage hidden tags to shapes or slides'"
    AddFeature "ShowFormSelectSlidesByTag", "Select slides by tag(s) or specific stamp", "ShowFormSelectSlidesByTag", "Advanced", "Instrumenta > Advanced > Inside splitbutton 'Manage hidden tags to shapes or slides'", "[Advanced] > Advanced > Inside splitbutton 'Manage hidden tags to shapes or slides'"
    AddFeature "ExcelMailMerge", "Mail merge this slide based on Excel-file (duplicate slide within this presentation)", "ExcelMailMerge", "Advanced", "Instrumenta > Advanced", "[Advanced] > Advanced"
    AddFeature "ExcelFullFileMailMerge", "Mail merge full presentation based on Excel-file (duplicate presentation file)", "ExcelFullFileMailMerge", "Advanced", "Instrumenta > Advanced", "[Advanced] > Advanced"
    AddFeature "InsertMergeField", "Insert empty merge field", "InsertMergeField", "Advanced", "Instrumenta > Advanced", "[Advanced] > Advanced"
    AddFeature "ImportHeadersFromExcel", "Insert merge fields from Excel-file", "ImportHeadersFromExcel", "Advanced", "Instrumenta > Advanced", "[Advanced] > Advanced"
    AddFeature "ManualMailMerge", "Manually replace all merge fields on all slides", "ManualMailMerge", "Advanced", "Instrumenta > Advanced", "[Advanced] > Advanced"
    AddFeature "ShowSlideGraderButton", "Slide Grader", "ShowSlideGrader", "Advanced", "Instrumenta > Advanced", "[Advanced] > Advanced"
    AddFeature "ShowFeatureSearchButton", "Find Instrumenta features", "ShowFeatureSearch", "Advanced", "Instrumenta > Advanced", "[Advanced] > Advanced"
    AddFeature "ShowSettingsDialogButton", "Instrumenta settings", "ShowSettings", "Advanced", "Instrumenta > Advanced", "[Advanced] > Advanced"
    AddFeature "ShowAboutDialogButton", "Show about dialog", "ShowAboutDialog", "Advanced", "Instrumenta > Advanced", "[Advanced] > Advanced"
    ReDim Preserve Features(1 To FeatureCount) As FeatureData
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading features: " & Err.Description, vbCritical
End Sub

Private Sub AddFeature(id As String, lbl As String, action As String, grp As String, tabSingle As String, tabMulti As String)
    FeatureCount = FeatureCount + 1
    Features(FeatureCount).id = id
    Features(FeatureCount).label = lbl
    Features(FeatureCount).OnAction = action
    Features(FeatureCount).GroupLabel = grp
    Features(FeatureCount).TabSingleView = tabSingle
    Features(FeatureCount).TabMultiView = tabMulti
End Sub


Public Function SearchFeatures(query As String) As String
    Dim results() As Long
    Dim resultCount As Long
    Dim i As Long
    Dim searchTerm As String
    
    searchTerm = LCase(Trim(query))
    ReDim results(1 To FeatureCount) As Long
    resultCount = 0
    
    If Len(searchTerm) = 0 Then
        For i = 1 To FeatureCount
            resultCount = resultCount + 1
            results(resultCount) = i
        Next i
    Else

        For i = 1 To FeatureCount
            If InStr(1, LCase(Features(i).label), searchTerm) > 0 Or _
               InStr(1, LCase(Features(i).GroupLabel), searchTerm) > 0 Or _
               InStr(1, LCase(Features(i).TabSingleView), searchTerm) > 0 Or _
               InStr(1, LCase(Features(i).TabMultiView), searchTerm) > 0 Then
                resultCount = resultCount + 1
                results(resultCount) = i
            End If
        Next i
    End If
    
    If resultCount = 0 Then
    Exit Function
    Else

    ReDim Preserve results(1 To resultCount) As Long
    End If
    
    Dim resultStr As String
    For i = 1 To resultCount
        resultStr = resultStr & results(i) & "|"
    Next i
    
    SearchFeatures = resultStr
End Function

Public Sub ExecuteFeature(actionName As String)
    On Error GoTo ErrorHandler
    
     If LCase$(right$(actionName, 9)) = "_onaction" Then
     MsgBox "This feature cannot be executed from this window, please use the Ribbon"
     Exit Sub
     End If
    
    Application.run actionName
    Exit Sub
ErrorHandler:
    MsgBox "Could not execute: " & actionName & vbCrLf & Err.Description, vbExclamation
End Sub

Public Function GetFeatureByIndex(index As Long) As FeatureData
    GetFeatureByIndex = Features(index)
End Function

Public Function GetFeatureCount() As Long
    GetFeatureCount = FeatureCount
End Function

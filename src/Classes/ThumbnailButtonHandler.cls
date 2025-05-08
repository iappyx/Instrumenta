VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThumbnailButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents ClickOverlay As MSForms.CommandButton
Attribute ClickOverlay.VB_VarHelpID = -1

Private Sub ClickOverlay_Click()
    Call InsertSlideLibrarySlide.ToggleCheckBox(ClickOverlay.Tag)
End Sub

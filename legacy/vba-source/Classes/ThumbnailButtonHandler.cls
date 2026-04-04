Attribute VB_Name = "ThumbnailButtonHandler"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public WithEvents ClickOverlay As MSForms.CommandButton
Attribute ClickOverlay.VB_VarHelpID = -1

Private Sub ClickOverlay_Click()
    Call InsertSlideLibrarySlide.ToggleCheckBox(ClickOverlay.Tag)
End Sub

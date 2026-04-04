Attribute VB_Name = "RibbonEventsClass"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private WithEvents app As Application
Attribute app.VB_VarHelpID = -1

Public Sub Initialize()
    Set app = Application
End Sub

Private Sub app_WindowSelectionChange(ByVal sel As Selection)
    If Not InstrumentaRibbon Is Nothing Then
        InstrumentaRibbon.Invalidate
    End If
End Sub


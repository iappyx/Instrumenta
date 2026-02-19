VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonEventsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


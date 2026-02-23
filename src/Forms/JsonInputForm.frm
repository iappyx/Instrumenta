Attribute VB_Name = "JsonInputForm"
Attribute VB_Base = "0{C4FDC70F-F1E6-4AC2-894B-73E4812B630C}{C9F2A657-1A60-4C83-A712-4D0E6BA52B6A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public result As String
Public Cancelled As Boolean

Private Sub cmdOK_Click()
    result = Me.txtJson.text
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub


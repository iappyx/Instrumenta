Attribute VB_Name = "JsonInputForm"
Attribute VB_Base = "0{E114A9DC-39B1-4A3E-A53C-0690F4887DD5}{5BDCD8F1-2D45-4695-A6DE-957709981F75}"
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


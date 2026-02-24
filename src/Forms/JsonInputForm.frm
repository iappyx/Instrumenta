Attribute VB_Name = "JsonInputForm"
Attribute VB_Base = "0{B702245E-2E00-421F-99C0-1697FFA45624}{6D7B21C3-34AA-482C-9E44-B53843DFDB0C}"
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


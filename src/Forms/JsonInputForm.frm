Attribute VB_Name = "JsonInputForm"
Attribute VB_Base = "0{8A63EFFB-9013-4EC1-A644-2F71B5A4FCE9}{824C1FF4-55D0-42F8-8C17-20C2599E7D92}"
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


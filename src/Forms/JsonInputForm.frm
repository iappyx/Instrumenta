Attribute VB_Name = "JsonInputForm"
Attribute VB_Base = "0{E6BDDE4F-5BD1-43F0-A61F-DE679965E638}{0258AD55-E545-4D3E-91EC-219AF08EEB78}"
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


Attribute VB_Name = "JsonInputForm"
Attribute VB_Base = "0{5BFD87AA-AC9E-471D-9334-4F6A574D8EA2}{D560CB50-BD52-4BCB-8A15-04555393DBBB}"
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


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim userName As String
    userName = Environ("USERNAME") ' Gépre bejelentkezett felhasználó

    lblWelcome.Caption = "Üdvözöllek, " & userName & "! Kérlek, válaszd ki a munkaterületed:"
End Sub

Private Sub cmdEBike_Click()
    MsgBox "eBike Warranty kiválasztva!"
End Sub

Private Sub cmdEM_Click()
    MsgBox "EM Warranty kiválasztva!"
End Sub

Private Sub cmdClean_Click()
    MsgBox "Cleanliness kiválasztva!"
End Sub

Private Sub cmdEMPBQZ_Click()
    MsgBox "EMPB/QZ kiválasztva!"
End Sub


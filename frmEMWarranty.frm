VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEMWarranty 
   Caption         =   "EM Warranty"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13755
   OleObjectBlob   =   "frmEMWarranty.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEMWarranty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblUser_Click()

End Sub

Private Sub UserForm_Click()

End Sub
Private Sub UserForm_Initialize()
    lblUser.Caption = "Felhasználó: " & Environ("USERNAME")
End Sub

Private Sub cmdReklamacio_Click()
    Me.Hide
    frmEM_Reklamacio.Show
    Me.Show
End Sub

Private Sub cmdEgyeb_Click()
    
    Me.Hide
    frmEM_Egyeb.Show
    Me.Show

End Sub


Private Sub cmdLeltar_Click()
    MsgBox "A 'Garancia leltár' modul még fejlesztés alatt.", vbInformation
End Sub

Private Sub cmdPS_Click()
    MsgBox "A 'Problem Solving' modul még nem véglegesített.", vbInformation
End Sub


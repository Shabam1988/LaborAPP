VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWelcome 
   Caption         =   "LaborApp � �dv�zl�k�perny�"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13755
   OleObjectBlob   =   "frmWelcome.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdEBike_Click()
    Me.Hide
    frmEBike.Show
    Me.Show
End Sub


Private Sub cmdEM_Click()
    Me.Hide
    frmEMWarranty.Show
    Me.Show
End Sub


Private Sub cmdClean_Click()
    MsgBox "Cleanliness kiv�lasztva!"
End Sub

Private Sub cmdEMPBQZ_Click()
    MsgBox "EMPB/QZ kiv�lasztva!"
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim userName As String
    userName = Environ("USERNAME") ' G�pre bejelentkezett felhaszn�l�

    lblWelcome.Caption = "�dv�z�llek, " & userName & "! K�rlek, v�laszd ki a munkater�leted:"
End Sub




VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEM_Egyeb 
   Caption         =   "EM Egy�b Idok Elsz�mol�sa"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13755
   OleObjectBlob   =   "frmEM_Egyeb.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEM_Egyeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim i As Integer
    Dim userName As String

    ' Felhaszn�l�n�v ki�r�sa
    userName = Environ("USERNAME")
    lblUser.Caption = "Felhaszn�l�: " & userName

    ' D�tumv�laszt� - �v
    For i = Year(Date) - 5 To Year(Date) + 5
        cmbYear.AddItem i
    Next i
    cmbYear.Value = Year(Date)

    ' D�tumv�laszt� - H�nap
    For i = 1 To 12
        cmbMonth.AddItem Format(DateSerial(2000, i, 1), "mmmm")
    Next i
    cmbMonth.ListIndex = Month(Date) - 1

    ' D�tumv�laszt� - Nap
    For i = 1 To 31
        cmbDay.AddItem i
    Next i
    cmbDay.Value = Day(Date)

    ' T�puslista felt�lt�se mind a 7 ComboBox-ba
    Dim tipusLista As Variant
    Dim cmb As MSForms.ComboBox

    tipusLista = Array( _
        "Teamboard", "Megbesz�l�sek", "TPM", "Bels� vizsg�latok", "Fuvarszervez�s", "Csomagkik�ld�s", _
        "Oktat�s", "Tr�ningek", "Garancia Level1", "Garancia Level2", "L�togat�s", "R�ntgen", _
        "EMPB", "QZ", "Rajzpecs�tel�s", "Jegyz�k�nyv k�sz�t�s", "Minta Logisztik�ja", "ELPC", _
        "Dokument�ci� kezel�s, friss�t�s", "Ebike Support", "eDate", "EdocPro", "PowerBI riport�l�s, k�sz�t�s, m�dos�t�s", _
        "QC nyit�s", "Labor fejleszt�s", "CRM selejtez�s", "Reklam�ci�k logisztikai kezel�se", _
        "Problem Solving", "eBIKE extra vizsg�lat", "DU el�k�sz�t�s, mos�s", "Szelekt�v teszt", _
        "Anyaghi�ny", "NVH Zajvizsg�lat", "Form LABS 3D", "Szabads�g")

    For i = 1 To 7
        Set cmb = Me.Controls("cmbEgyebTipus" & i)
        Dim t As Variant
        For Each t In tipusLista
            cmb.AddItem t
        Next t
    Next i

    ' Sz�ml�l� friss�t�se indul�skor
    Call FrissitIdoSzamlalo
End Sub

' D�tum �ssze�ll�t�sa
Function GetSelectedDate() As String
    GetSelectedDate = Format(DateSerial(cmbYear.Value, cmbMonth.ListIndex + 1, cmbDay.Value), "yyyy.mm.dd")
End Function

' Napok friss�t�se h�nap/�v v�ltoz�skor
Private Sub cmbMonth_Change(): RefreshDays: End Sub
Private Sub cmbYear_Change(): RefreshDays: End Sub

Private Sub RefreshDays()
    Dim numDays As Integer, selYear As Integer, selMonth As Integer
    Dim prevDay As Integer, i As Integer

    If cmbYear.Value = "" Or cmbMonth.ListIndex = -1 Then Exit Sub

    selYear = cmbYear.Value
    selMonth = cmbMonth.ListIndex + 1
    prevDay = val(cmbDay.Value)

    numDays = Day(DateSerial(selYear, selMonth + 1, 0))
    cmbDay.Clear
    For i = 1 To numDays
        cmbDay.AddItem i
    Next i

    If prevDay <= numDays Then
        cmbDay.Value = prevDay
    Else
        cmbDay.Value = numDays
    End If
End Sub

' Val�s idej� visszasz�ml�l�s
Sub FrissitIdoSzamlalo()
    Dim totalTime As Long
    Dim i As Integer, val As String
    Dim fennmarado As Long
    Const teljesIdo As Long = 460 ' 8 �r�s munkarend percben

    totalTime = 0
    For i = 1 To 7
        val = Me.Controls("txtEgyebIdo" & i).Text
        If IsNumeric(val) Then totalTime = totalTime + CLng(val)
    Next i

    fennmarado = teljesIdo - totalTime
    If fennmarado < 0 Then fennmarado = 0

    lblOsszeg.Caption = "Marad�k id�: " & fennmarado & " perc"
End Sub

' Change esem�nyek a 7 mez�re
Private Sub txtEgyebIdo1_Change(): FrissitIdoSzamlalo: End Sub
Private Sub txtEgyebIdo2_Change(): FrissitIdoSzamlalo: End Sub
Private Sub txtEgyebIdo3_Change(): FrissitIdoSzamlalo: End Sub
Private Sub txtEgyebIdo4_Change(): FrissitIdoSzamlalo: End Sub
Private Sub txtEgyebIdo5_Change(): FrissitIdoSzamlalo: End Sub
Private Sub txtEgyebIdo6_Change(): FrissitIdoSzamlalo: End Sub
Private Sub txtEgyebIdo7_Change(): FrissitIdoSzamlalo: End Sub

' Seg�df�ggv�ny: valid�ci� sz�mra
Function IsInputValid() As Boolean
    Dim i As Integer, val As String

    For i = 1 To 7
        val = Trim(Me.Controls("txtEgyebIdo" & i).Text)
        If val <> "" And Not IsNumeric(val) Then
            IsInputValid = False
            Exit Function
        End If
    Next i
    IsInputValid = True
End Function

' Seg�df�ggv�ny: mez�k t�rl�se
Sub ClearFields()
    Dim i As Integer
    For i = 1 To 7
        Me.Controls("cmbEgyebTipus" & i).Value = ""
        Me.Controls("txtEgyebIdo" & i).Text = ""
    Next i
End Sub

Private Sub cmdModositas_Click()
    If Not IsInputValid() Then
        MsgBox "Csak sz�mot adhatsz meg az id� mez�kbe!", vbExclamation
        Exit Sub
    End If

    Dim wb As Workbook, ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim datum As String, user As String

    datum = GetSelectedDate()
    user = Environ("USERNAME")

    Set wb = Workbooks.Open("\\Mc-file04\qas$\Laboratory\Project\LaborAPP\LaborDB.xlsx")
    Set ws = wb.Sheets("EgyebIdok")

    ' T�rl�s el�z� bejegyz�sek
    For i = ws.Cells(ws.Rows.Count, 1).End(xlUp).row To 2 Step -1
        If ws.Cells(i, 1).Value = datum And ws.Cells(i, 2).Value = user Then
            ws.Rows(i).Delete
        End If
    Next i

    ' �jrament�s
    For i = 1 To 7
        If Me.Controls("cmbEgyebTipus" & i).Value <> "" And _
           IsNumeric(Me.Controls("txtEgyebIdo" & i).Text) Then

            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
            ws.Cells(lastRow, 1).Value = datum
            ws.Cells(lastRow, 2).Value = user
            ws.Cells(lastRow, 3).Value = Me.Controls("cmbEgyebTipus" & i).Value
            ws.Cells(lastRow, 4).Value = CLng(Me.Controls("txtEgyebIdo" & i).Text)
        End If
    Next i

    wb.Close True
    MsgBox "Sikeres m�dos�t�s!", vbInformation
    ClearFields
    FrissitIdoSzamlalo
End Sub

Private Sub cmdBetoltes_Click()
    Dim wb As Workbook, ws As Worksheet
    Dim datum As String, user As String
    Dim i As Long, sorSzam As Integer

    datum = GetSelectedDate()
    user = Environ("USERNAME")
    sorSzam = 1

    Set wb = Workbooks.Open("\\Mc-file04\qas$\Laboratory\Project\LaborAPP\LaborDB.xlsx")
    Set ws = wb.Sheets("EgyebIdok")

    ClearFields

    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        If ws.Cells(i, 1).Value = datum And ws.Cells(i, 2).Value = user Then
            If sorSzam <= 7 Then
                Me.Controls("cmbEgyebTipus" & sorSzam).Value = ws.Cells(i, 3).Value
                Me.Controls("txtEgyebIdo" & sorSzam).Text = ws.Cells(i, 4).Value
                sorSzam = sorSzam + 1
            End If
        End If
    Next i

    wb.Close False
    FrissitIdoSzamlalo
End Sub

Private Sub cmdMentes_Click()
    If Not IsInputValid() Then
        MsgBox "Csak sz�mot adhatsz meg az id� mez�kbe!", vbExclamation
        Exit Sub
    End If

    Dim wb As Workbook, ws As Worksheet
    Dim i As Integer, datum As String, user As String
    Dim tipus As String, ido As String, lastRow As Long

    Set wb = Workbooks.Open("\\Mc-file04\qas$\Laboratory\Project\LaborAPP\LaborDB.xlsx")
    Set ws = wb.Sheets("EgyebIdok")

    datum = GetSelectedDate()
    user = Environ("USERNAME")

    ' Ellen�rz�s: van-e m�r adat?
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        If ws.Cells(i, 1).Value = datum And ws.Cells(i, 2).Value = user Then
            MsgBox "Erre a napra m�r van adatod! Haszn�ld a M�dos�t�s funkci�t.", vbExclamation
            wb.Close False
            Exit Sub
        End If
    Next i

    ' Ment�s
    For i = 1 To 7
        tipus = Me.Controls("cmbEgyebTipus" & i).Value
        ido = Me.Controls("txtEgyebIdo" & i).Text

        If tipus <> "" And IsNumeric(ido) Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
            ws.Cells(lastRow, 1).Value = datum
            ws.Cells(lastRow, 2).Value = user
            ws.Cells(lastRow, 3).Value = tipus
            ws.Cells(lastRow, 4).Value = CLng(ido)
        End If
    Next i

    wb.Close True
    MsgBox "Sikeres ment�s!", vbInformation
    ClearFields
    FrissitIdoSzamlalo
End Sub



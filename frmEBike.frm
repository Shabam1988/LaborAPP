VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEBike 
   Caption         =   "eBike – Adatrögzítés"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13755
   OleObjectBlob   =   "frmEBike.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEBike"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim i As Integer
    Dim j As Integer
    Dim userName As String
    Dim tipusLista As Variant

    ' Felhasználó név
    userName = Environ("USERNAME")
    Me.lblUserName.Caption = userName

    ' Év ComboBox
    For i = Year(Date) - 5 To Year(Date) + 5
        Me.cmbYear.AddItem i
    Next i
    Me.cmbYear.Value = Year(Date)

    ' Hónap ComboBox
    For i = 1 To 12
        Me.cmbMonth.AddItem Format(DateSerial(2000, i, 1), "mmmm")
    Next i
    Me.cmbMonth.ListIndex = Month(Date) - 1

    ' Nap ComboBox
    For i = 1 To 31
        Me.cmbDay.AddItem i
    Next i
    Me.cmbDay.Value = Day(Date)

    ' Munkarend ComboBox
    cmbMunkarend.AddItem "8 óra"
    cmbMunkarend.AddItem "12 óra"
    cmbMunkarend.ListIndex = 0

    ' Egyéb típuslista feltöltése mind a 7 ComboBoxba
    tipusLista = Array("Teamboard", "Megbeszélések", "TPM", "Belsõ vizsgálatok", "Fuvarszervezés", _
                       "Csomagkiküldés", "Oktatás", "Tréningek", "Garancia Level1", "Garancia Level2", _
                       "Látogatás", "Röntgen", "EMPB", "QZ", "Rajzpecsételés", "Jegyzõkönyv készítés", _
                       "Minta Logisztikája", "ELPC", "Dokumentáció kezelés, frissítés", "Ebike Support", _
                       "eDate", "EdocPro", "PowerBI riportálás, készítés, módosítás", "QC nyitás", _
                       "Labor fejlesztés", "CRM selejtezés", "Reklamációk logisztikai kezelése", _
                       "Problem Solving", "eBIKE extra vizsgálat", "DU elõkészítés, mosás", _
                       "Szelektív teszt", "Anyaghiány", "NVH Zajvizsgálat", "Form LABS 3D", "Szabadság")

    For i = 1 To 7
        With Me.Controls("cmbEgyebTipus" & i)
            For j = LBound(tipusLista) To UBound(tipusLista)
                .AddItem tipusLista(j)
            Next j
        End With
    Next i
End Sub

Private Sub cmbMonth_Change()
    Call RefreshDays
End Sub

Private Sub cmbYear_Change()
    Call RefreshDays
End Sub

Private Sub RefreshDays()
    Dim numDays As Integer
    Dim selYear As Integer
    Dim selMonth As Integer
    Dim i As Integer
    Dim prevDay As Integer

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

Function GetSelectedDate() As String
    GetSelectedDate = Format(DateSerial(cmbYear.Value, cmbMonth.ListIndex + 1, cmbDay.Value), "yyyy.mm.dd")
End Function

Sub FrissitIdoSzamlalo()
    Dim darabokOsszesen As Long
    Dim i As Integer
    Dim vizsgalatiIdo As Long
    Dim munkarendIdo As Long
    Dim fennmarado As Long
    Dim egyebIdo As Long

    ' Vizsgálatok összes ideje
        darabokOsszesen = 0
    If IsNumeric(txtBDU38.Text) Then darabokOsszesen = darabokOsszesen + val(txtBDU38.Text)
    If IsNumeric(txtGEN3.Text) Then darabokOsszesen = darabokOsszesen + val(txtGEN3.Text)
    If IsNumeric(txtGEN4.Text) Then darabokOsszesen = darabokOsszesen + val(txtGEN4.Text)
    If IsNumeric(txtGEN4_BES3.Text) Then darabokOsszesen = darabokOsszesen + val(txtGEN4_BES3.Text)
    If IsNumeric(txtGEN3_BES3.Text) Then darabokOsszesen = darabokOsszesen + val(txtGEN3_BES3.Text)
    If IsNumeric(txtBDU31.Text) Then darabokOsszesen = darabokOsszesen + val(txtBDU31.Text)
    If IsNumeric(txtBDU34.Text) Then darabokOsszesen = darabokOsszesen + val(txtBDU34.Text)


    vizsgalatiIdo = darabokOsszesen * 30

    ' Egyéb idõk összesítése
    egyebIdo = 0
    For i = 1 To 7
        If Trim(Me.Controls("cmbEgyebTipus" & i).Text) <> "" And IsNumeric(Me.Controls("txtEgyebIdo" & i).Text) Then
            egyebIdo = egyebIdo + val(Me.Controls("txtEgyebIdo" & i).Text)
        End If
    Next i

    ' Munkarend kiválasztása
    Select Case cmbMunkarend.Value
        Case "8 óra": munkarendIdo = 460
        Case "12 óra": munkarendIdo = 675
        Case Else: munkarendIdo = 0
    End Select

    ' Maradék idõ kiszámítása
    fennmarado = munkarendIdo - vizsgalatiIdo - egyebIdo
    If fennmarado < 0 Then fennmarado = 0

    lblIdoSzamlalo.Caption = "Maradék idõ: " & fennmarado & " perc"
End Sub

Function IsNumericInputValid() As Boolean
    Dim inputs As Variant
    Dim val As Variant
    Dim i As Integer

    inputs = Array(txtBDU38.Text, txtGEN3.Text, txtGEN4.Text, txtGEN4_BES3.Text, txtGEN3_BES3.Text, txtBDU31.Text, txtBDU34.Text)

    For i = LBound(inputs) To UBound(inputs)
        val = Trim(inputs(i))
        If val = "" Then
        ElseIf Not IsNumeric(val) Then
            IsNumericInputValid = False
            Exit Function
        End If
    Next i

    ' Egyéb idõk is számok legyenek, ha meg vannak adva
    For i = 1 To 7
        If Trim(Me.Controls("txtEgyebIdo" & i).Text) <> "" Then
            If Not IsNumeric(Me.Controls("txtEgyebIdo" & i).Text) Then
                IsNumericInputValid = False
                Exit Function
            End If
        End If
    Next i

    IsNumericInputValid = True
End Function

Sub ClearInputFields()
    txtBDU38.Text = ""
    txtGEN3.Text = ""
    txtGEN4.Text = ""
    txtGEN4_BES3.Text = ""
    txtGEN3_BES3.Text = ""
    txtBDU31.Text = ""
    txtBDU34.Text = ""

    For i = 1 To 7
        Me.Controls("cmbEgyebTipus" & i).Text = ""
        Me.Controls("txtEgyebIdo" & i).Text = ""
    Next i
End Sub

Private Sub cmbMunkarend_Change()
    Call FrissitIdoSzamlalo
End Sub

' --- 1. Típus darabszám mezõk Change eseményei ---
Private Sub txtBDU38_Change(): Call FrissitIdoSzamlalo: End Sub
Private Sub txtGEN3_Change(): Call FrissitIdoSzamlalo: End Sub
Private Sub txtGEN4_Change(): Call FrissitIdoSzamlalo: End Sub
Private Sub txtGEN4_BES3_Change(): Call FrissitIdoSzamlalo: End Sub
Private Sub txtGEN3_BES3_Change(): Call FrissitIdoSzamlalo: End Sub
Private Sub txtBDU31_Change(): Call FrissitIdoSzamlalo: End Sub
Private Sub txtBDU34_Change(): Call FrissitIdoSzamlalo: End Sub

' --- 2. Egyéb idõk mezõk Change eseményei ---
Private Sub txtEgyebIdo1_Change(): Call FrissitIdoSzamlalo: End Sub
Private Sub txtEgyebIdo2_Change(): Call FrissitIdoSzamlalo: End Sub
Private Sub txtEgyebIdo3_Change(): Call FrissitIdoSzamlalo: End Sub
Private Sub txtEgyebIdo4_Change(): Call FrissitIdoSzamlalo: End Sub
Private Sub txtEgyebIdo5_Change(): Call FrissitIdoSzamlalo: End Sub
Private Sub txtEgyebIdo6_Change(): Call FrissitIdoSzamlalo: End Sub
Private Sub txtEgyebIdo7_Change(): Call FrissitIdoSzamlalo: End Sub

' --- 3. Mentés ---
Private Sub cmdSave_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim dbPath As String
    Dim lastRow As Long
    Dim i As Long
    Dim user As String
    Dim entryDate As String
    Dim exists As Boolean
    Dim openedHere As Boolean

    dbPath = "\\Mc-file04\qas$\Laboratory\Project\LaborAPP\LaborDB.xlsx"
    user = Environ("USERNAME")
    entryDate = GetSelectedDate()

    If Not IsNumericInputValid() Then
        MsgBox "Kérlek, csak számokat adj meg a mezõkben!", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wb = Workbooks("LaborDB.xlsx")
    If wb Is Nothing Then
        Set wb = Workbooks.Open(dbPath)
        openedHere = True
    End If
    On Error GoTo 0

    Set ws = wb.Sheets("eBike")
    Set ws2 = wb.Sheets("EgyebIdok")

    ' eBike mentés
    exists = False
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        If ws.Cells(i, 1).Value = entryDate And ws.Cells(i, 2).Value = user Then
            exists = True
            Exit For
        End If
    Next i

    If exists Then
        MsgBox "Erre a napra már van adatod! Használd a 'Betöltés' és 'Módosítás' lehetõséget.", vbExclamation
        If openedHere Then wb.Close SaveChanges:=False
        Set wb = Nothing
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    ws.Cells(lastRow, 1).Value = entryDate
    ws.Cells(lastRow, 2).Value = user
    ws.Cells(lastRow, 3).Value = val(txtBDU38.Text)
    ws.Cells(lastRow, 4).Value = val(txtGEN3.Text)
    ws.Cells(lastRow, 5).Value = val(txtGEN4.Text)
    ws.Cells(lastRow, 6).Value = val(txtGEN4_BES3.Text)
    ws.Cells(lastRow, 7).Value = val(txtGEN3_BES3.Text)
    ws.Cells(lastRow, 8).Value = val(txtBDU31.Text)
    ws.Cells(lastRow, 9).Value = val(txtBDU34.Text)

    ' EgyebIdok mentés
    For i = 1 To 7
        If Trim(Me.Controls("cmbEgyebTipus" & i).Text) <> "" And Trim(Me.Controls("txtEgyebIdo" & i).Text) <> "" Then
            lastRow = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).row + 1
            ws2.Cells(lastRow, 1).Value = entryDate
            ws2.Cells(lastRow, 2).Value = user
            ws2.Cells(lastRow, 3).Value = Me.Controls("cmbEgyebTipus" & i).Text
            ws2.Cells(lastRow, 4).Value = val(Me.Controls("txtEgyebIdo" & i).Text)
        End If
    Next i

    MsgBox "Adatok sikeresen elmentve!", vbInformation
    If openedHere Then wb.Close SaveChanges:=True
    Set wb = Nothing
    Call ClearInputFields
End Sub

' --- 4. Betöltés ---
Private Sub cmdLoad_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim dbPath As String
    Dim user As String
    Dim entryDate As String
    Dim i As Long
    Dim k As Long
    Dim openedHere As Boolean

    dbPath = "\\Mc-file04\qas$\Laboratory\Project\LaborAPP\LaborDB.xlsx"
    user = Environ("USERNAME")
    entryDate = GetSelectedDate()

    On Error Resume Next
    Set wb = Workbooks("LaborDB.xlsx")
    If wb Is Nothing Then
        Set wb = Workbooks.Open(dbPath)
        openedHere = True
    End If
    On Error GoTo 0

    Set ws = wb.Sheets("eBike")
    Set ws2 = wb.Sheets("EgyebIdok")

    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        If ws.Cells(i, 1).Value = entryDate And ws.Cells(i, 2).Value = user Then
            txtBDU38.Text = ws.Cells(i, 3).Value
            txtGEN3.Text = ws.Cells(i, 4).Value
            txtGEN4.Text = ws.Cells(i, 5).Value
            txtGEN4_BES3.Text = ws.Cells(i, 6).Value
            txtGEN3_BES3.Text = ws.Cells(i, 7).Value
            txtBDU31.Text = ws.Cells(i, 8).Value
            txtBDU34.Text = ws.Cells(i, 9).Value
            Exit For
        End If
    Next i

    ' EgyebIdok betöltése max 7 sorig
    k = 1
    For i = 2 To ws2.Cells(ws2.Rows.Count, "A").End(xlUp).row
        If ws2.Cells(i, 1).Value = entryDate And ws2.Cells(i, 2).Value = user Then
            If k <= 7 Then
                Me.Controls("cmbEgyebTipus" & k).Text = ws2.Cells(i, 3).Value
                Me.Controls("txtEgyebIdo" & k).Text = ws2.Cells(i, 4).Value
                k = k + 1
            End If
        End If
    Next i

    If openedHere Then wb.Close SaveChanges:=False
    Set wb = Nothing
End Sub

' --- 5. Módosítás ---
Private Sub cmdUpdate_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim dbPath As String
    Dim user As String
    Dim entryDate As String
    Dim i As Long
    Dim updated As Boolean
    Dim openedHere As Boolean

    dbPath = "\\Mc-file04\qas$\Laboratory\Project\LaborAPP\LaborDB.xlsx"
    user = Environ("USERNAME")
    entryDate = GetSelectedDate()

    If Not IsNumericInputValid() Then
        MsgBox "Kérlek, csak számokat adj meg a mezõkben!", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wb = Workbooks("LaborDB.xlsx")
    If wb Is Nothing Then
        Set wb = Workbooks.Open(dbPath)
        openedHere = True
    End If
    On Error GoTo 0

    Set ws = wb.Sheets("eBike")
    Set ws2 = wb.Sheets("EgyebIdok")

    ' eBike módosítása
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        If ws.Cells(i, 1).Value = entryDate And ws.Cells(i, 2).Value = user Then
            ws.Cells(i, 3).Value = val(txtBDU38.Text)
            ws.Cells(i, 4).Value = val(txtGEN3.Text)
            ws.Cells(i, 5).Value = val(txtGEN4.Text)
            ws.Cells(i, 6).Value = val(txtGEN4_BES3.Text)
            ws.Cells(i, 7).Value = val(txtGEN3_BES3.Text)
            ws.Cells(i, 8).Value = val(txtBDU31.Text)
            ws.Cells(i, 9).Value = val(txtBDU34.Text)
            Exit For
        End If
    Next i

    ' EgyebIdok korábbi sorainak törlése
    For i = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).row To 2 Step -1
        If ws2.Cells(i, 1).Value = entryDate And ws2.Cells(i, 2).Value = user Then
            ws2.Rows(i).Delete
        End If
    Next i

    ' EgyebIdok újraszúrás
    For i = 1 To 7
        If Trim(Me.Controls("cmbEgyebTipus" & i).Text) <> "" And Trim(Me.Controls("txtEgyebIdo" & i).Text) <> "" Then
            lastRow = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).row + 1
            ws2.Cells(lastRow, 1).Value = entryDate
            ws2.Cells(lastRow, 2).Value = user
            ws2.Cells(lastRow, 3).Value = Me.Controls("cmbEgyebTipus" & i).Text
            ws2.Cells(lastRow, 4).Value = val(Me.Controls("txtEgyebIdo" & i).Text)
        End If
    Next i

    MsgBox "Adatok frissítve!", vbInformation
    If openedHere Then wb.Close SaveChanges:=True
    Set wb = Nothing
End Sub



VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEM_Reklamacio 
   Caption         =   "EM Reklam�ci� kezel�s"
   ClientHeight    =   13410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21765
   OleObjectBlob   =   "frmEM_Reklamacio.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEM_Reklamacio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' Felhaszn�l�n�v be�ll�t�sa
    lblUser.Caption = "Felhaszn�l�: " & Environ("USERNAME")

    ' D�tumv�laszt� felt�lt�se
    Dim i As Integer
    For i = Year(Date) - 5 To Year(Date) + 5
        cmbYear.AddItem i
    Next i
    cmbYear.Value = Year(Date)

    For i = 1 To 12
        cmbMonth.AddItem Format(DateSerial(2000, i, 1), "mmmm")
    Next i
    cmbMonth.ListIndex = Month(Date) - 1

    For i = 1 To 31
        cmbDay.AddItem i
    Next i
    cmbDay.Value = Day(Date)

    ' �ra �s perc mez�k
    txtIdo.Text = Format(Time, "hh:mm")
    
    ' Term�kcsoport �rt�kek be�ll�t�sa
    cmbTermekcsoport.Clear
    cmbTermekcsoport.AddItem "Airmax"
    cmbTermekcsoport.AddItem "BPA"
    cmbTermekcsoport.AddItem "DPO"
    cmbTermekcsoport.AddItem "ECAS2"
    cmbTermekcsoport.AddItem "ECF"
    cmbTermekcsoport.AddItem "ECM"
    cmbTermekcsoport.AddItem "ECo2"
    cmbTermekcsoport.AddItem "FED"
    cmbTermekcsoport.AddItem "GBB"
    cmbTermekcsoport.AddItem "GBM"

    ' Field vagy 0km �rt�kek be�ll�t�sa
    cmbField0km.Clear
    cmbField0km.AddItem "Field"
    cmbField0km.AddItem "0km"
    
        ' Vev� �rt�kek be�ll�t�sa
    cmbVevo.Clear
    Dim customers As Variant
    customers = Array( _
        "Valeo Rodach", "Valeo Rakovnik", "Valeo Nogent", "Valeo Korea", _
        "Valeo Martorelles", "Valeo Mioveni", "Valeo Ninjing", "Valeo Chihuahua", _
        "Valeo Zaragoza", "Valeo Uitenhage", "Valeo Bursa (TR)", "Valeo Togliatti", _
        "Valeo Titu", "Mahle Kirchberg", "Mahle Czhech", "Mahle Rouffach", _
        "Mahle South Afrika", "Mahle Neustadt", "Mahle Mnichovo", "Mahle Shenyang", _
        "Mahle Senica", "Mahle Korea", "Mahle Spain", "Hanon Charleville", _
        "Hanon Ilava", "Hanon Turkey", "Visteon USA", "HVCC USA", _
        "Air International", "Air International SK", "Denso Italy", "Denso CZ", _
        "Denso UK", "Wirthwein Polska", "RBCC", "Bosch B�hl", _
        "RBPL-Ostrow", "RBPL-Mirkow", "RBKB-Korea", "VW AG", _
        "Air International Shanghai", "Hanon Gebze", "Bosch Energy and Body Systems", _
        "Bosch PSA e transmissions", "INA-SCHAEFFLER KG", "Daejung Europe" _
    )

    Dim j As Integer
    For j = LBound(customers) To UBound(customers)
        cmbVevo.AddItem customers(j)
    Next j
End Sub
Private Sub cmbYear_Change()
    Call RefreshDays
End Sub
Private Sub cmbMonth_Change()
    Call RefreshDays
End Sub
Private Sub RefreshDays()
    If cmbYear.Value = "" Or cmbMonth.ListIndex = -1 Then Exit Sub

    Dim selYear As Integer: selYear = cmbYear.Value
    Dim selMonth As Integer: selMonth = cmbMonth.ListIndex + 1
    Dim prevDay As Integer: prevDay = val(cmbDay.Value)
    Dim numDays As Integer: numDays = Day(DateSerial(selYear, selMonth + 1, 0))
    Dim i As Integer

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
Function GetDatumIdo() As String
    On Error Resume Next
    Dim datum As Date
    datum = DateSerial(cmbYear.Value, cmbMonth.ListIndex + 1, cmbDay.Value)
    GetDatumIdo = Format(datum, "yyyy.mm.dd")
End Function
Sub UpdateStatusLabel()
    Dim lezaras As String
    Dim level2open As String
    Dim level2close As String
    
    lezaras = Trim(txtLezarasDatuma.Text)
    level2open = Trim(txtLevel2NyitasDatuma.Text)
    level2close = Trim(txtLevel2LezarasDatuma.Text)
    
    If lezaras = "" Then
        lblStatus.Caption = "Nyitott"
    ElseIf lezaras <> "" And level2open = "" Then
        lblStatus.Caption = "Lez�rt"
    ElseIf level2open <> "" And level2close = "" Then
        lblStatus.Caption = "Level2 folyamatban"
    ElseIf level2close <> "" Then
        lblStatus.Caption = "Level2 lez�rt"
    Else
        lblStatus.Caption = "Nyitott" ' Default fallback
    End If
End Sub

Private Sub cmdMentes_Click()
    Dim qcszam As String, cikkszam As String, termekcsoport As String, mennyiseg As String
    Dim field0km As String, vevo As String, hibaleiras As String, ecuszam As String
    Dim vevoiReklamaciosSzam As String, folderPath As String, mappaEleres As String
    Dim baseFolderPath As String, yearFolderPath As String, yearMonthFolderPath As String
    Dim datumIdo As String, felhasznalo As String
    Dim wb As Workbook, ws As Worksheet, nextRow As Long

    ' Aktu�lis id�adatok
    datumIdo = GetDatumIdo()
    felhasznalo = Environ("USERNAME")
    
    ' Mez�k beolvas�sa
    qcszam = IIf(txtQC.Text = "", "23000", txtQC.Text)
    cikkszam = Replace(txtCikkszam.Text, " ", "")
    termekcsoport = cmbTermekcsoport.Text
    mennyiseg = txtMennyiseg.Text
    field0km = cmbField0km.Text
    vevo = cmbVevo.Text
    hibaleiras = txtHibaLeiras.Text
    ecuszam = txtEcuSzam.Text
    vevoiReklamaciosSzam = txtVevoiReklamaciosSzam.Text

    ' ?? Mappagener�l�s
    Select Case termekcsoport
        Case "Airmax": baseFolderPath = "\\mc-file04\ED_QMM_MC$\QMM8\Garantieanalyse\ED\Airmax\"
        Case "BPA": baseFolderPath = "\\mc-file04\ED_QMM_MC$\QMM8\Garantieanalyse\ED\BPA\"
        Case "DPO": baseFolderPath = "\\mc-file04\ED_QMM_MC$\QMM8\Garantieanalyse\ED\DPO\"
        Case "ECAS2": baseFolderPath = "\\mc-file04\ED_QMM_MC$\QMM8\Garantieanalyse\ED\ECAS2\"
        Case "ECF": baseFolderPath = "\\mc-file04\ED_QMM_MC$\QMM8\Garantieanalyse\ED\ECF\"
        Case "ECM": baseFolderPath = "\\mc-file04\ED_QMM_MC$\QMM8\Garantieanalyse\ED\ECM\"
        Case "ECo2": baseFolderPath = "\\mc-file04\ED_QMM_MC$\QMM8\Garantieanalyse\ED\ECo2\"
        Case "FED": baseFolderPath = "\\mc-file04\ED_QMM_MC$\QMM8\Garantieanalyse\ED\FED\"
        Case "GBB": baseFolderPath = "\\mc-file04\ED_QMM_MC$\QMM8\Garantieanalyse\ED\GBB\"
        Case "GBM": baseFolderPath = "\\mc-file04\ED_QMM_MC$\QMM8\Garantieanalyse\ED\GBM\"
        Case Else
            MsgBox "Ismeretlen term�kcsoport!", vbExclamation
            Exit Sub
    End Select
    
    Dim currentYear As String: currentYear = Format(Now, "yyyy")
    Dim currentYearMonth As String: currentYearMonth = Format(Now, "yyyy.mm")
    yearFolderPath = baseFolderPath & currentYear
    yearMonthFolderPath = yearFolderPath & "\" & currentYearMonth

    If Dir(yearFolderPath, vbDirectory) = "" Then MkDir yearFolderPath
    If Dir(yearMonthFolderPath, vbDirectory) = "" Then MkDir yearMonthFolderPath

    folderPath = yearMonthFolderPath & "\" & qcszam & "_" & cikkszam & "_" & field0km
    If ecuszam <> "" Then folderPath = folderPath & "_" & ecuszam
    folderPath = folderPath & "_" & hibaleiras & "_" & Format(Now, "yyyymmdd")
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
    mappaEleres = folderPath

    ' ?? Adatok ment�se az adatb�zisba
    Set wb = Workbooks.Open("\\Mc-file04\qas$\Laboratory\Project\LaborAPP\LaborDB.xlsx")
    Set ws = wb.Sheets("EMWarranty")
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1

    ' Mez�k �r�sa
    ws.Cells(nextRow, 1).Value = datumIdo
    ws.Cells(nextRow, 2).Value = txtIdo.Text
    ws.Cells(nextRow, 3).Value = felhasznalo
    ws.Cells(nextRow, 4).Value = qcszam
    ws.Cells(nextRow, 5).Value = cikkszam
    ws.Cells(nextRow, 6).Value = termekcsoport
    ws.Cells(nextRow, 7).Value = mennyiseg
    ws.Cells(nextRow, 8).Value = field0km
    ws.Cells(nextRow, 9).Value = vevo
    ws.Cells(nextRow, 10).Value = hibaleiras
    ws.Cells(nextRow, 11).Value = ecuszam
    ws.Cells(nextRow, 12).Value = vevoiReklamaciosSzam
    ws.Cells(nextRow, 13).Value = mappaEleres
    ' 14�20 mez�k: �resen hagyva, majd m�dos�t�skor t�lthet�k (lez�r�s, Level2 nyit�s, stb.)

    wb.Save
    wb.Close
    
    ' ?? Email k�ld�s megk�rdez�se
    If MsgBox("Szeretn�l emailt k�ldeni err�l a reklam�ci�r�l?", vbYesNo + vbQuestion, "Email k�ld�s") = vbYes Then
        Call GenerateEmail
    End If

    MsgBox "Reklam�ci� sikeresen mentve!", vbInformation
    Unload Me
End Sub

Sub GenerateEmail()
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim emailBody As String
    Dim subject As String
    Dim folderPath As String

    ' El�r�si �tvonal a txtMappaEleres mez�b�l
    folderPath = txtMappaEleres.Text

    ' Email t�rgy
    subject = "Be�rkezett " & cmbTermekcsoport.Text & " reklam�ci� r�szletei " & Format(Now, "yyyy.mm.dd")

    ' Email t�rzs � HTML form�z�ssal
    emailBody = "<p style='font-size:14pt;'>Sziasztok,</p>" & _
                "<p>A mai napon (" & Format(Now, "yyyy.mm.dd") & ") be�rkezett reklam�ci� r�szletei:</p>" & _
                "<table border='1' style='border-collapse:collapse; width: 100%;'>" & _
                "<tr><th>QC sz�m</th><th>Cikksz�m</th><th>Term�kcsoport</th><th>Mennyis�g</th><th>Field vagy 0km</th><th>Vev�</th><th>Hibale�r�s</th><th>ECU sz�m</th><th>Vev�i reklam�ci�s sz�m</th></tr>" & _
                "<tr><td>" & txtQC.Text & "</td><td>" & txtCikkszam.Text & "</td><td>" & cmbTermekcsoport.Text & "</td><td>" & txtMennyiseg.Text & "</td><td>" & cmbField0km.Text & "</td><td>" & cmbVevo.Text & "</td><td>" & txtHibaLeiras.Text & "</td><td>" & txtEcuSzam.Text & "</td><td>" & txtVevoiReklamaciosSzam.Text & "</td></tr></table>" & _
                "<p>El�r�si �tvonal: <a href='" & folderPath & "'>" & folderPath & "</a></p>" & _
                "<p>QC nyit�shoz k�rem kit�lteni:</p>"

    emailBody = emailBody & "<table border='1' style='border-collapse:collapse; width: 100%;'>" & _
                "<tr><td style='width: 50%;'>Reklam�ci� t�pusa / Complaint mode (0km / Field):</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>Statisztik�s / In statistic (Igen / Nem):</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>Cikksz�m / Bosch part number:</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>Vev�i cikksz�m / Customer part number:</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>Mikor tal�lt�k a darabot / Repair date:</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>Aut�ba �p�t�s d�tuma / Registration date:</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>Gy�rt�si d�tum / Manufacturing date:</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>VIN:</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>Megtett km / Mileage:</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>Honnan �rkezett / From which location:</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>Hiba le�r�sa / Claim description:</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>Mennyis�g / Quantity:</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>Sorsz�m/ Folder number:</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>Vev�i reklam�ci�s sz�m / Customer claim:</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>�tfut�si id� / Elv�gz�si d�tum/ End date (24/48/ vagy t�bb):</td><td style='width: 50%;'></td></tr>" & _
                "<tr><td style='width: 50%;'>Required end:</td><td style='width: 50%;'></td></tr>" & _
                "</table>"

    ' Outlook p�ld�ny l�trehoz�sa
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0)

    ' Email be�ll�t�sa
    With outlookMail
        .To = "" ' C�mzettek k�s�bb ker�lnek be�ll�t�sra manu�lisan
        .cc = ""
        .BCC = ""
        .subject = subject
        .HTMLBody = emailBody
        .Display
    End With
End Sub

Private Sub cmdKereses_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim path As String
    path = "\\Mc-file04\qas$\Laboratory\Project\LaborAPP\LaborDB.xlsx"

    On Error Resume Next
    Set wb = Workbooks("LaborDB.xlsx")
    If wb Is Nothing Then
        Set wb = Workbooks.Open(path, ReadOnly:=True)
    End If
    On Error GoTo 0

    Set ws = wb.Sheets("EMWarranty")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' Keres�si mez�k
    Dim keresettMezok(1 To 10) As String
    keresettMezok(1) = Trim(txtQC.Text)
    keresettMezok(2) = Trim(txtCikkszam.Text)
    keresettMezok(3) = Trim(cmbTermekcsoport.Text)
    keresettMezok(4) = Trim(txtMennyiseg.Text)
    keresettMezok(5) = Trim(cmbField0km.Text)
    keresettMezok(6) = Trim(cmbVevo.Text)
    keresettMezok(7) = Trim(txtHibaLeiras.Text)
    keresettMezok(8) = Trim(txtEcuSzam.Text)
    keresettMezok(9) = Trim(txtVevoiReklamaciosSzam.Text)
    keresettMezok(10) = Trim(txtMegjegyzes.Text)

    Dim i As Long, j As Long
    Dim egyezes As Boolean
    Dim talalatokSzama As Long: talalatokSzama = 0

    ' Els� k�r: tal�latsz�ml�l�s
    For i = 2 To lastRow
        egyezes = True
        If keresettMezok(1) <> "" And InStr(ws.Cells(i, 4).Text, keresettMezok(1)) = 0 Then egyezes = False
        If keresettMezok(2) <> "" And InStr(ws.Cells(i, 5).Text, keresettMezok(2)) = 0 Then egyezes = False
        If keresettMezok(3) <> "" And InStr(ws.Cells(i, 6).Text, keresettMezok(3)) = 0 Then egyezes = False
        If keresettMezok(4) <> "" And InStr(ws.Cells(i, 7).Text, keresettMezok(4)) = 0 Then egyezes = False
        If keresettMezok(5) <> "" And InStr(ws.Cells(i, 8).Text, keresettMezok(5)) = 0 Then egyezes = False
        If keresettMezok(6) <> "" And InStr(ws.Cells(i, 9).Text, keresettMezok(6)) = 0 Then egyezes = False
        If keresettMezok(7) <> "" And InStr(ws.Cells(i, 10).Text, keresettMezok(7)) = 0 Then egyezes = False
        If keresettMezok(8) <> "" And InStr(ws.Cells(i, 11).Text, keresettMezok(8)) = 0 Then egyezes = False
        If keresettMezok(9) <> "" And InStr(ws.Cells(i, 12).Text, keresettMezok(9)) = 0 Then egyezes = False
        If keresettMezok(10) <> "" And InStr(ws.Cells(i, 20).Text, keresettMezok(10)) = 0 Then egyezes = False
        If egyezes Then talalatokSzama = talalatokSzama + 1
    Next i

    If talalatokSzama = 0 Then
        MsgBox "Nincs tal�lat a keresett felt�telek alapj�n!", vbInformation
        lstTalalatok.Clear
        wb.Close False
        Exit Sub
    End If

    ' M�sodik k�r: t�mb felt�lt�se
    Dim talalatok() As Variant
    ReDim talalatok(0 To talalatokSzama, 0 To 19)

    ' Fejl�c sor
    For j = 0 To 19
        talalatok(0, j) = ws.Cells(1, j + 1).Text
    Next j

    Dim talalatIndex As Long: talalatIndex = 0

    For i = 2 To lastRow
        egyezes = True
        If keresettMezok(1) <> "" And InStr(ws.Cells(i, 4).Text, keresettMezok(1)) = 0 Then egyezes = False
        If keresettMezok(2) <> "" And InStr(ws.Cells(i, 5).Text, keresettMezok(2)) = 0 Then egyezes = False
        If keresettMezok(3) <> "" And InStr(ws.Cells(i, 6).Text, keresettMezok(3)) = 0 Then egyezes = False
        If keresettMezok(4) <> "" And InStr(ws.Cells(i, 7).Text, keresettMezok(4)) = 0 Then egyezes = False
        If keresettMezok(5) <> "" And InStr(ws.Cells(i, 8).Text, keresettMezok(5)) = 0 Then egyezes = False
        If keresettMezok(6) <> "" And InStr(ws.Cells(i, 9).Text, keresettMezok(6)) = 0 Then egyezes = False
        If keresettMezok(7) <> "" And InStr(ws.Cells(i, 10).Text, keresettMezok(7)) = 0 Then egyezes = False
        If keresettMezok(8) <> "" And InStr(ws.Cells(i, 11).Text, keresettMezok(8)) = 0 Then egyezes = False
        If keresettMezok(9) <> "" And InStr(ws.Cells(i, 12).Text, keresettMezok(9)) = 0 Then egyezes = False
        If keresettMezok(10) <> "" And InStr(ws.Cells(i, 20).Text, keresettMezok(10)) = 0 Then egyezes = False

        If egyezes Then
            talalatIndex = talalatIndex + 1
            For j = 0 To 19
                talalatok(talalatIndex, j) = ws.Cells(i, j + 1).Text
            Next j
        End If
    Next i

    ' ListBox bet�lt�s
    With lstTalalatok
        .ColumnCount = 20
        .ColumnHeads = True
        .List = talalatok
    End With

    wb.Close False
End Sub





Function GetStatusFromRow(ws As Worksheet, rowNum As Long) As String
    Dim lezaras As String, level2open As String, level2close As String
    lezaras = Trim(ws.Cells(rowNum, 14).Value)
    level2open = Trim(ws.Cells(rowNum, 19).Value)
    level2close = Trim(ws.Cells(rowNum, 20).Value)

    If lezaras = "" Then
        GetStatusFromRow = "Nyitott"
    ElseIf lezaras <> "" And level2open = "" Then
        GetStatusFromRow = "Lez�rt"
    ElseIf level2open <> "" And level2close = "" Then
        GetStatusFromRow = "Level2 folyamatban"
    ElseIf level2close <> "" Then
        GetStatusFromRow = "Level2 lez�rt"
    Else
        GetStatusFromRow = "Nyitott"
    End If
End Function
Private Sub lstTalalatok_Click()
    If lstTalalatok.ListIndex < 0 Then Exit Sub

    ' Sorindex ment�se
    Dim row As Integer
    row = lstTalalatok.ListIndex

    ' === D�tum �s id� sz�tv�laszt�sa ===
    Dim teljesDatumIdo As String
    Dim datumResz As String
    Dim idoResz As String

    teljesDatumIdo = lstTalalatok.List(row, 0)

    If InStr(teljesDatumIdo, " ") > 0 Then
        datumResz = Split(teljesDatumIdo, " ")(0)
        idoResz = Split(teljesDatumIdo, " ")(1)
    Else
        datumResz = teljesDatumIdo
        idoResz = ""
    End If

    ' D�tum sz�tszed�se mez�kre
    If datumResz <> "" Then
        On Error Resume Next ' Ha a form�tum hib�s, ne dobjon hib�t
        cmbYear.Value = Split(datumResz, ".")(0)
        cmbMonth.Value = CInt(Split(datumResz, ".")(1))
        cmbDay.Value = CInt(Split(datumResz, ".")(2))
        On Error GoTo 0
    End If

    ' Id� visszat�lt�s
    txtIdo.Text = idoResz

    ' Egy�b mez�k visszat�lt�se
    lblUser.Caption = "Felhaszn�l�: " & lstTalalatok.List(row, 2)
    txtQC.Text = lstTalalatok.List(row, 3)
    txtCikkszam.Text = lstTalalatok.List(row, 4)
    cmbTermekcsoport.Text = lstTalalatok.List(row, 5)
    txtMennyiseg.Text = lstTalalatok.List(row, 6)
    cmbField0km.Text = lstTalalatok.List(row, 7)
    cmbVevo.Text = lstTalalatok.List(row, 8)
    txtHibaLeiras.Text = lstTalalatok.List(row, 9)
    txtEcuSzam.Text = lstTalalatok.List(row, 10)
    txtVevoiReklamaciosSzam.Text = lstTalalatok.List(row, 11)
    txtMappaEleres.Text = lstTalalatok.List(row, 12)
    txtLezarasDatuma.Text = lstTalalatok.List(row, 13)
    txtLezarta.Text = lstTalalatok.List(row, 14)
    txtLevel2NyitasDatuma.Text = lstTalalatok.List(row, 15)
    txtLevel2LezarasDatuma.Text = lstTalalatok.List(row, 16)
    txtLevel2Lezarta.Text = lstTalalatok.List(row, 17)
    lblStatus.Caption = lstTalalatok.List(row, 18)
    txtMegjegyzes.Text = lstTalalatok.List(row, 19)

    ' St�tusz friss�t�se
    Call UpdateStatusLabel
End Sub
Private Sub cmdModositas_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim regiMappaNev As String, ujMappaNev As String
    Dim datumResz As String, idoResz As String
    Dim fs As Object

    ' === Megnyitjuk az adatb�zist ===
    Set wb = Workbooks.Open("\\Mc-file04\qas$\Laboratory\Project\LaborAPP\LaborDB.xlsx")
    Set ws = wb.Sheets("EMWarranty")

    ' === Sorok sz�ma ===
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' === Mappa el�r�si �t azonos�t�k�nt ===
    regiMappaNev = Trim(txtMappaEleres.Text)

   
    ' === V�gigmegy�nk a sorokon, mappan�v alapj�n ===
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 13).Value) = regiMappaNev Then

            ' === Mez�k friss�t�se ===
            ws.Cells(i, 1).Value = datumResz
            ws.Cells(i, 2).Value = txtIdo.Text
            ws.Cells(i, 3).Value = Mid(lblUser.Caption, 15) ' �Felhaszn�l�: ��
            ws.Cells(i, 4).Value = txtQC.Text
            ws.Cells(i, 5).Value = txtCikkszam.Text
            ws.Cells(i, 6).Value = cmbTermekcsoport.Text
            ws.Cells(i, 7).Value = txtMennyiseg.Text
            ws.Cells(i, 8).Value = cmbField0km.Text
            ws.Cells(i, 9).Value = cmbVevo.Text
            ws.Cells(i, 10).Value = txtHibaLeiras.Text
            ws.Cells(i, 11).Value = txtEcuSzam.Text
            ws.Cells(i, 12).Value = txtVevoiReklamaciosSzam.Text

            ' === �j mappa�tvonal gener�l�sa ===
            ujMappaNev = "\\Mc-file04\qas$\Laboratory\Project\LaborAPP\" & _
                cmbTermekcsoport.Text & "\" & Replace(datumResz, ".", "-") & _
                "_" + txtQCSzam.Text & "_" & txtCikkszam.Text & "_" & _
                Replace(txtHibaLeiras.Text, " ", "_")

            ' === Ha megv�ltozott, akkor mappa �tnevez�s ===
            If regiMappaNev <> ujMappaNev Then
                Set fs = CreateObject("Scripting.FileSystemObject")
                If fs.FolderExists(regiMappaNev) Then
                    On Error Resume Next
                    fs.MoveFolder regiMappaNev, ujMappaNev
                    On Error GoTo 0
                End If
            End If

            ' === Friss�tj�k az el�r�si �t mez�t ===
            ws.Cells(i, 13).Value = ujMappaNev

            ' === Kil�p�s, ha megtal�ltuk a sort ===
            Exit For
        End If
    Next i

    wb.Close SaveChanges:=True
    MsgBox "A m�dos�t�s sikeresen megt�rt�nt!", vbInformation
End Sub



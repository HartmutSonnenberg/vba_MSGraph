Attribute VB_Name = "modPrognose_Veroeffentlichen"
'===============================================================================
' Name:         modPrognose_Veroeffentlichen
' Purpose:      Modul bietet Funktionen an, um die Inhalte der Sheets "KF_Prognose_24_25"
'               bis "KF_Prognose_24_25" in vorgegebene Verzeichnisse/Dateien (staticdata!)
'               als csv oder xlsx Dateien via Email zu veröffentlichen
' Remarks:      Author H. Sonnenberg im Juni 2024
' Functions:    Fill_Workbook, Daten_Formatieren, Hole_Template, Create_Kunden_CSV,
'               Get_Pfadangaben, Create_Syneco_CSV
' Properties:
' Methods:      Delete_Worksheets, Prognose_Veroeffentlichen
' Started:      20.06.2024
' Modified:     09.01.2026  Emailversand über MSGraph
'===============================================================================
Option Explicit

'===============================================================================
' Name: Function Prognose_Veroeffentlichen
' Input:
'   None
' Output:
'   None
' Purpose:  exportiert die Inhalte der Sheets "KF_Prognose_24_25" bis
'           "KF_Prognose_24_25" in vorgegebene Verzeichnisse/Dateien (staticdata)
'           und startet den Emailversand über MSGraph
' Remarks:  Juli 2024 H. Sonnenberg, geändert 2026/01
'===============================================================================
Public Sub Prognose_Veroeffentlichen()
    Dim arrTemplates() As String                'Array mit den Namen, Pfaden und Erweiterungen aller Templates (Staticdata Spalte E, D und F)
    Dim arrOutpSyn() As String                  'Array mit den Namen, Ausgabepfaden, Erweiterungen, verwendetes Templates, Worksheet, Datumsvariante? für alle Outputdateien Syneco
    Dim arrOutpKunden() As String               'Array mit den Namen, Ausgabepfaden, Erweiterungen, verwendetes Templates, Worksheet, Datumsvariante? für alle Outputdateien für Kunden
    
    Dim firstRowBereich As Long                 'wird mit den je ersten zeilen der 3 Bereiche Templates, Output Syneco und Output Kunden bestückt
    Dim lastRowBereich As Long                  'wird mit den je letzten zeilen der 3 Bereiche Templates, Output Syneco und Output Kunden bestückt
    Dim row As Long                             'Schleifenvariable
    Dim wbExport As Workbook
    
    Dim foreCastFile As String                  'Speicherort und Name der Dateikopie am Ende der Veröffentlichung
    Dim fehler As String
    Dim pfadAngabe As String
    Dim exitCode As Boolean
    
    On Error GoTo ErrHandler
    
    'SOF-883: Bei Fehlermeldungen - diese in eigenem Formular anzeigen. Kurzfrist PFC Versand kann abgebrochen werden
    If CheckMeldungen = False Then
        fehler = HoleFehlermeldungen
        ShowErrorsAtForm fehler
        exitCode = frmCheckErrors.ExitProcess
    End If
    If exitCode Then Exit Sub
    
    fehler = vbNullString
    
    StatusBar "Dateien der Kurzfrist PFC Freigabe werden erzeugt"
    
    Boost_VBA True
    
    '1. die 3 wichtigen Bereichnsnamen rngStartTemplates, rngStartOutputSyneco und rngStartOutputKunden prüfen
    If wsStaticData.Range("rngStartTemplates") Is Nothing Then _
            Err.Raise 901, "Prognose_Veroeffentlichen", "die benannte Zelle ""rngStartTemplates"" (in B11) konnte nicht gefunden werden!"
    If wsStaticData.Range("rngStartOutputSyneco") Is Nothing Then _
            Err.Raise 901, "Prognose_Veroeffentlichen", "die benannte Zelle ""rngStartOutputSyneco"" (in B16) konnte nicht gefunden werden!"
    If wsStaticData.Range("rngStartOutputKunden") Is Nothing Then _
            Err.Raise 901, "Prognose_Veroeffentlichen", "die benannte Zelle ""rngStartOutputKunden"" (eigentlich in B26) konnte nicht gefunden werden!"
    
    '2. Namen und Pfade der Templates in Array arrTemplates() speichern - mit 3 Spalten
    firstRowBereich = wsStaticData.Range("rngStartTemplates").row
    lastRowBereich = wsStaticData.Range("rngStartTemplates").Offset(0, 1).End(xlDown).row
    
    '*******************************************************************************************************
    '************* die Templates im Array arrTemplates() mit Name, Pfad, Erweiterung speichern *************
    '*******************************************************************************************************
    
    ReDim arrTemplates(lastRowBereich - firstRowBereich, 2)
    ' arrTemplates(row, 1) enthält den Pfad zum Template: auf Backslash prüfen und Verzeichnis Laufwerk prüfen
    For row = 0 To lastRowBereich - firstRowBereich
        'Dateinamen in Spalte E
        arrTemplates(row, 0) = wsStaticData.Cells(wsStaticData.Range("rngStartTemplates").Offset(row, 0).row, _
            wsStaticData.Range("rngStartTemplates").Offset(row, 3).Column).value 'Name
        'Pfade stehen je nach Version in den Spalten D, J oder Q
        If Get_Pfadangaben("templates", row, pfadAngabe, fehler) = True Then
            arrTemplates(row, 1) = CheckPathName(pfadAngabe)
        Else
            Err.Raise 902, "Prognose_Veroeffentlichen", "Templatepfad konnte nicht bestimmt werden. " & vbCrLf & fehler
        End If
        If CheckPath(arrTemplates(row, 1)) = False Then _
            Err.Raise 902, "Prognose_Veroeffentlichen", "Template-Verzeichnis: """ & arrTemplates(row, 1) & """ nicht gefunden!"
        
        'Erweiterung in Spalte F
        arrTemplates(row, 2) = wsStaticData.Cells(wsStaticData.Range("rngStartTemplates").Offset(row, 0).row, wsStaticData.Range("rngStartTemplates").Offset(row, 4).Column).value 'Erweiterung
    Next row
    
    '********************************************************************************************************************************************
    '***********  den Bereich Output für die Syneco festhalten und im Array arrOutpSyn() speichern
    '***********  arrOutpSyn() hat Spalten Namen, Ausgabepfaden, Erweiterungen, verwendetes Templates, Worksheet-Name, Datumsvariante  ''''''''''
    '********************************************************************************************************************************************
    
    '3. alle Infos für den Output Syneco in  arrOutpSyn() speichern - mit den Spalten Namen, Ausgabepfaden, Erweiterungen, verwendetes Templates, Worksheet-Name, Datumsvariante?
    firstRowBereich = wsStaticData.Range("rngStartOutputSyneco").row
    lastRowBereich = wsStaticData.Range("rngStartOutputSyneco").Offset(0, 1).End(xlDown).row
    
    ReDim arrOutpSyn(lastRowBereich - firstRowBereich, 5)
    ' arrOutpSyn(row, 1) enthält den Ausgabepfad: auf Backslash prüfen und Verzeichnis Laufwerk prüfen
    For row = 0 To lastRowBereich - firstRowBereich
        'Dateinamen in Spalte E
        arrOutpSyn(row, 0) = wsStaticData.Cells(wsStaticData.Range("rngStartOutputSyneco").Offset(row, 0).row, wsStaticData.Range("rngStartOutputSyneco").Offset(row, 4).Column).value 'Name
        'Pfade stehen je nach Version in den Spalten D, J oder Q
        If Get_Pfadangaben("outputsyneco", row, pfadAngabe, fehler) = True Then
            arrOutpSyn(row, 1) = CheckPathName(pfadAngabe)
        Else
            Err.Raise 902, "Prognose_Veroeffentlichen", "Pfad für den syneco Output konnte nicht bestimmt werden. " & vbCrLf & fehler
        End If
        If CheckPath(arrOutpSyn(row, 1)) = False Then _
            Err.Raise 902, "Prognose_Veroeffentlichen", "Syneco Output-Verzeichnis: """ & arrOutpSyn(row, 1) & """ nicht gefunden!"
        
        arrOutpSyn(row, 2) = wsStaticData.Cells(wsStaticData.Range("rngStartOutputSyneco").Offset(row, 0).row, _
            wsStaticData.Range("rngStartOutputSyneco").Offset(row, 5).Column).value 'Erweiterung
        arrOutpSyn(row, 3) = wsStaticData.Cells(wsStaticData.Range("rngStartOutputSyneco").Offset(row, 0).row, _
            wsStaticData.Range("rngStartOutputSyneco").Offset(row, 2).Column).value 'verwendetes Templates
        arrOutpSyn(row, 4) = wsStaticData.Cells(wsStaticData.Range("rngStartOutputSyneco").Offset(row, 0).row, _
            wsStaticData.Range("rngStartOutputSyneco").Offset(row, 6).Column).value 'Worksheet-Name
        arrOutpSyn(row, 5) = wsStaticData.Cells(wsStaticData.Range("rngStartOutputSyneco").Offset(row, 0).row, _
            wsStaticData.Range("rngStartOutputSyneco").Offset(row, 7).Column).value 'mit Datumsvariante?
    Next row
    'in Output Syneco gibt es nur zwei Varianten: alles Exceldateien (xlsx) - 2 auf der Basis eines Templates, alle anderen ohne Templates und CSV Dateien!
    
    
    '********************************************************************************************************************************************
    '***********  den Bereich Output für Kunden festhalten und im Array arrOutpKunden() speichern
    '***********  arrOutpSyn() hat Spalten Namen, Ausgabepfaden, Erweiterungen, verwendetes Templates, Worksheet-Name, Datumsvariante  ''''''''''
    '********************************************************************************************************************************************
    
    
    '4. alle Infos für den Output für Kunden in  rngStartOutputKunden() speichern - mit den Spalten Namen, Ausgabepfaden, Erweiterungen, verwendetes Templates, Worksheet, Datumsvariante?
    firstRowBereich = wsStaticData.Range("rngStartOutputKunden").row
    lastRowBereich = wsStaticData.Range("rngStartOutputKunden").Offset(0, 1).End(xlDown).row
    
    ReDim arrOutpKunden(lastRowBereich - firstRowBereich, 5)
    ' arrOutpKunden(row, 1) enthält den Ausgabepfad: auf Backslash prüfen und Verzeichnis Laufwerk prüfen
    For row = 0 To lastRowBereich - firstRowBereich
        'csv- Dateien
        If LCase(wsStaticData.Cells(wsStaticData.Range("rngStartOutputKunden").Offset(row, 0).row, wsStaticData.Range("rngStartOutputKunden").Offset(row, 5).Column).value) = "csv" Then
            'Dateinamen in Spalte E
            arrOutpKunden(row, 0) = wsStaticData.Cells(wsStaticData.Range("rngStartOutputKunden").Offset(row, 0).row, wsStaticData.Range("rngStartOutputKunden").Offset(row, 4).Column).value 'Name
            'Pfade stehen je nach Version in den Spalten D, J oder Q
            If Get_Pfadangaben("outputkunden", row, pfadAngabe, fehler) = True Then
                arrOutpKunden(row, 1) = CheckPathName(pfadAngabe)
            Else
                Err.Raise 902, "Prognose_Veroeffentlichen", "Pfad für den syneco Output konnte nicht bestimmt werden. " & vbCrLf & fehler
            End If
            If CheckPath(arrOutpKunden(row, 1)) = False Then _
                Err.Raise 902, "Prognose_Veroeffentlichen", "Kunden Output-Verzeichnis: """ & arrOutpKunden(row, 1) & """ nicht gefunden!"
        
'            If InStr(LCase(ThisWorkbook.Name), "entw") > 0 Or InStr(LCase(ThisWorkbook.Name), "test") > 0 Then
'                arrOutpKunden(row, 1) = CheckPathName(wsStaticData.Cells(wsStaticData.Range("rngStartOutputKunden").Offset(row, 0).row, wsStaticData.Range("rngStartOutputKunden").Offset(row, 8).Column).value) 'Ausgabepfad
'            Else
'                arrOutpKunden(row, 1) = CheckPathName(wsStaticData.Cells(wsStaticData.Range("rngStartOutputKunden").Offset(row, 0).row, wsStaticData.Range("rngStartOutputKunden").Offset(row, 3).Column).value) 'Ausgabepfad produktiv
'            End If
            If InStr(LCase(ThisWorkbook.Name), "entw") = 0 Then _
                If CheckPath(arrOutpKunden(row, 1)) = False Then Err.Raise 902, "Prognose_Veroeffentlichen", "Kunden Output-Verzeichnis: """ & arrOutpKunden(row, 1) & """ nicht gefunden!"
            arrOutpKunden(row, 2) = wsStaticData.Cells(wsStaticData.Range("rngStartOutputKunden").Offset(row, 0).row, _
                wsStaticData.Range("rngStartOutputKunden").Offset(row, 5).Column).value 'Erweiterung
            arrOutpKunden(row, 3) = "na"
            arrOutpKunden(row, 4) = wsStaticData.Cells(wsStaticData.Range("rngStartOutputKunden").Offset(row, 0).row, _
                wsStaticData.Range("rngStartOutputKunden").Offset(row, 6).Column).value 'Worksheet-Name
            arrOutpKunden(row, 5) = "na"
        Else
            'Excel - xlsx-Dateien
            'Dateinamen in Spalte E
            arrOutpKunden(row, 0) = wsStaticData.Cells(wsStaticData.Range("rngStartOutputKunden").Offset(row, 0).row, wsStaticData.Range("rngStartOutputKunden").Offset(row, 4).Column).value 'Name
            'Pfade stehen je nach Version in den Spalten D, J oder Q
            If Get_Pfadangaben("outputkunden", row, pfadAngabe, fehler) = True Then
                arrOutpKunden(row, 1) = CheckPathName(pfadAngabe)
            Else
                Err.Raise 902, "Prognose_Veroeffentlichen", "Pfad für den Kunden Output konnte nicht bestimmt werden. " & vbCrLf & fehler
            End If
            
'            If InStr(LCase(ThisWorkbook.Name), "entw") > 0 Or InStr(LCase(ThisWorkbook.Name), "test") > 0 Then
'                arrOutpKunden(row, 1) = CheckPathName(wsStaticData.Cells(wsStaticData.Range("rngStartOutputKunden").Offset(row, 0).row, wsStaticData.Range("rngStartOutputKunden").Offset(row, 8).Column).value) 'Ausgabepfad
'            Else
'                arrOutpKunden(row, 1) = CheckPathName(wsStaticData.Cells(wsStaticData.Range("rngStartOutputKunden").Offset(row, 0).row, wsStaticData.Range("rngStartOutputKunden").Offset(row, 3).Column).value) 'Ausgabepfad produktiv
'            End If
            If CheckPath(arrOutpKunden(row, 1)) = False Then _
                Err.Raise 902, "Prognose_Veroeffentlichen", "Kunden Output-Verzeichnis: """ & arrOutpSyn(row, 1) & """ nicht gefunden!"
            
            arrOutpKunden(row, 2) = wsStaticData.Cells(wsStaticData.Range("rngStartOutputKunden").Offset(row, 0).row, wsStaticData.Range("rngStartOutputKunden").Offset(row, 5).Column).value 'Erweiterung
            arrOutpKunden(row, 3) = wsStaticData.Cells(wsStaticData.Range("rngStartOutputKunden").Offset(row, 0).row, wsStaticData.Range("rngStartOutputKunden").Offset(row, 2).Column).value 'verwendetes Templates
            arrOutpKunden(row, 4) = wsStaticData.Cells(wsStaticData.Range("rngStartOutputKunden").Offset(row, 0).row, wsStaticData.Range("rngStartOutputKunden").Offset(row, 6).Column).value 'Worksheet-Name
            arrOutpKunden(row, 5) = wsStaticData.Cells(wsStaticData.Range("rngStartOutputKunden").Offset(row, 0).row, wsStaticData.Range("rngStartOutputKunden").Offset(row, 7).Column).value 'mit Datumsvariante?
        End If
    Next row
    'in Output Kunden gibt es nur zwei Varianten: nur Exceldateien (xlsx) - auf der Basis eines Templates oder ohne!
    'ab 09/2025 SOF-1095 auch csv Dateien!
    
    
    '5. Output Syneco :arrOutpSyn() - mit den Spalten Namen, Ausgabepfaden, Erweiterungen, verwendetes Templates, Worksheet-Name, Datumsvariante?
    For row = LBound(arrOutpSyn(), 1) To UBound(arrOutpSyn(), 1)
        If LCase(arrOutpSyn(row, 2)) = "csv" Then
            'Dateiname:=arrOutpSyn(row, 1) & arrOutpSyn(row, 0) & "." & arrOutpSyn(row, 2)
            StatusBar arrOutpSyn(row, 1) & arrOutpSyn(row, 0) & "." & arrOutpSyn(row, 2) & " wird erzeugt und gespeichert!"
            If Create_Syneco_CSV(arrOutpSyn(row, 1) & arrOutpSyn(row, 0) & "." & arrOutpSyn(row, 2), ThisWorkbook.Worksheets(arrOutpSyn(row, 4)), fehler) = False Then _
                Err.Raise 901, "Prognose_Veroeffentlichen", fehler
        Else
            If arrOutpSyn(row, 3) <> vbNullString Then
                Set wbExport = Hole_Template(arrTemplates(), arrOutpSyn(row, 3), fehler)
                If fehler <> vbNullString Then Err.Raise 901, "Prognose_Veroeffentlichen", fehler
            Else
                Set wbExport = Application.Workbooks.Add
            End If
            
            StatusBar arrOutpSyn(row, 1) & arrOutpSyn(row, 0) & "." & arrOutpSyn(row, 2) & " wird erzeugt und gespeichert!"
            If Fill_Workbook(ThisWorkbook, wbExport, arrOutpSyn(row, 1), arrOutpSyn(row, 0), arrOutpSyn(row, 2), arrOutpSyn(row, 4), arrOutpSyn(row, 5), fehler) = False Then _
                    Err.Raise 901, "Prognose_Veroeffentlichen", fehler
            
            If Not wbExport Is Nothing Then wbExport.Close False
            Set wbExport = Nothing
        End If
    Next row
    
    
    '6. Output Kunden :arrOutpKunden() - mit den Spalten Namen, Ausgabepfaden, Erweiterungen, verwendetes Templates, Worksheet-Name, Datumsvariante?
    For row = LBound(arrOutpKunden(), 1) To UBound(arrOutpKunden(), 1)
        'im Augenblick gibt es keine CSV Dateien für Kunden....
        If LCase(arrOutpKunden(row, 2)) = "csv" Then
            'Dateiname:=arrOutpSyn(row, 1) & arrOutpSyn(row, 0) & "." & arrOutpSyn(row, 2)
            StatusBar arrOutpKunden(row, 0) & "." & arrOutpKunden(row, 2) & " wird erzeugt und gespeichert!"
            If Create_Kunden_CSV(arrOutpKunden(row, 1) & arrOutpKunden(row, 0) & "." & arrOutpKunden(row, 2), ThisWorkbook.Worksheets(arrOutpKunden(row, 4)), fehler) = False Then _
                Err.Raise 901, "Prognose_Veroeffentlichen", fehler
        Else
            If InStr(LCase(ThisWorkbook.Name), "entw") = 0 Then
                If arrOutpKunden(row, 3) <> vbNullString Then
                    Set wbExport = Hole_Template(arrTemplates(), arrOutpKunden(row, 3), fehler)
                    If fehler <> vbNullString Then Err.Raise 901, "Prognose_Veroeffentlichen", fehler
                Else
                    Set wbExport = Application.Workbooks.Add
                End If
                
                StatusBar arrOutpKunden(row, 1) & arrOutpKunden(row, 0) & "." & arrOutpKunden(row, 2) & " wird erzeugt und gespeichert!"
                If Fill_Workbook(ThisWorkbook, wbExport, arrOutpKunden(row, 1), arrOutpKunden(row, 0), arrOutpKunden(row, 2), arrOutpKunden(row, 4), arrOutpKunden(row, 5), fehler) = False Then _
                        Err.Raise 901, "Prognose_Veroeffentlichen", fehler
                
                If Not wbExport Is Nothing Then wbExport.Close False
                Set wbExport = Nothing
            End If
        End If
    Next row
    
    StatusBar "Dateikopie ohne VBA Code wird erstellt..."
    
    'SOF-883
    'Archivierung des Tools ohne Code als XLSX im Format YYYYMMTT_HHNN_Kurzfristgenerator
    If ConvertToXLSX(fehler) = False Then Err.Raise 901, "Prognose Veroeffentlichen", fehler
    
    'SOF-883
    If Hole_PrognoseDatei_FuerKunden(foreCastFile, fehler) = False Then Err.Raise 901, "Prognose Veroeffentlichen", fehler
    
    If Main_Create_Email_For_Customer(foreCastFile, fehler) = False Then Err.Raise 901, "Prognose Veroeffentlichen", fehler
    
    'SOF-1095 Teil 2: Anpassung Button "Kurzfrist PFC freigeben"
    'den Inhalt der Worksheets "ShortTerm_Forecast_15min_csv" und  "ShortTerm_PFC_15min_csv" exportieren
    'Pfade siehe staticdata! Zeile 37 ff. Achtung weitere Outputs sollen in den Zeilen 41 - 43 möglich sein
    
    ThisWorkbook.Activate

    'Buttons wieder färben
    Button_Dyeing wsControl.btnKurzfristPFC_Freigeben, en_Green

    MsgBox "alle Dateien wurden erfolgreich abgespeichert und Email an den Kunden versendet.", vbInformation, ThisWorkbook.Name
    
Ende:
    Boost_VBA False
    StatusBar_Reset
    
    Exit Sub
    
ErrHandler:
    Button_Dyeing wsControl.btnKurzfristPFC_Freigeben, en_Red
    If Not wbExport Is Nothing Then wbExport.Close False
    Set wbExport = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnPrognose_Click of VBA Dokument wsKurzfristEingabe", vbCritical, "EYE"
    Resume Ende
    
End Sub


'===============================================================================
' Name: Function Fill_Workbook
' Input:
'   ByRef wbQuelle As Workbook              -       dieses Workbook für Copy Aktion benötigt
'   ByRef wbExport As Workbook              -       die neue Exceldatei für Paste Aktion benötigt
'   ByVal ausgabePfad As String             -       Speicherort
'   ByVal dateiName As String               -       Dateiname
'   ByVal erweiterung As String             -       Erweiterung
'   ByVal wsQuellName As String             -       Name des Worksheets aus Quelle, dessen Daten kopiert werden sollen
'   ByVal Datumsvariante As Boolean         -       soll eine weitere Variant mit Datumsangabe im Dateinamen erzeugt werden
'   ByRef fehler As String                  -       Fehler werden per referenz an Caller zurückgegeben
' Output:
'   True / False
' Purpose:
' Remarks: Author H. Sonnenberg Juli 2024
Private Function Fill_Workbook(wbQuelle As Workbook, wbExport As Workbook, ByVal ausgabePfad As String, _
                                ByVal dateiName As String, ByVal erweiterung As String, _
                                ByVal wsQuellName As String, ByVal Datumsvariante As String, _
                                ByRef fehler As String) As Boolean

    Dim wsExport As Worksheet
    Dim fehlerCopyPaste As String
    Dim wsName  As String
    
    On Error GoTo ErrHandler
    
    Set wsExport = wbExport.Worksheets(1)
    
    If InStr(LCase(dateiName), "prognose") > 0 Then wsExport.Name = "Spot_Prognose" Else wsExport.Name = "KF_PFC"
   
    
    If Daten_Kopieren(wbQuelle.Worksheets(wsQuellName), wbQuelle.Worksheets(wsQuellName).Range("A1:E1000"), fehlerCopyPaste) = False Then _
            Err.Raise 901, "Fill_Workbook - Daten kopieren", fehlerCopyPaste
    
    If Daten_EinFuegen(wsExport.Range("A1:E1000"), True, fehlerCopyPaste) = False Then _
            Err.Raise 901, "Fill_Workbook - Daten Einfügen", fehlerCopyPaste
            
    Delete_Worksheets wbExport
    
    
    modCommon.File_Delete fileName:=ausgabePfad & dateiName & "." & erweiterung
    wbExport.SaveAs fileName:=ausgabePfad & dateiName & "." & erweiterung, FileFormat:=xlWorkbookDefault
    
    If Datumsvariante = "Ja" Then
        dateiName = dateiName & Format(Now(), "YYYYMMDD_HHMM")
        wbExport.SaveAs fileName:=ausgabePfad & dateiName & "." & erweiterung, FileFormat:=xlWorkbookDefault
    End If
    
    
    Fill_Workbook = True

Ende:
    
    Set wsExport = Nothing

    Exit Function
ErrHandler:
    fehler = Err.Description
    Resume Ende
End Function

'===============================================================================
' Name: Function Daten_Formatieren
' Input:
'   ByRef ws As Worksheet
' Output:
'   true / false
' Purpose: formatiert die Datumswerte in den Exportdateien - vorerst nicht verwendet
'       der Templates wegen
' Remarks: Author H. Sonnenberg Juli 2024
'===============================================================================
Private Function Daten_Formatieren(ws As Worksheet, fehler As String) As Boolean
    Dim lastRow As Long
    
    On Error GoTo ErrHandler
    
    ws.Range("B8").NumberFormat = "dd.mm.yyyy hh:mm"
    ws.Range("B11").NumberFormat = "dd.mm.yyyy hh:mm"

    ws.Range("A41:A761").NumberFormat = "ddd dd.mm.yyyy hh"
    
    Daten_Formatieren = True
    Exit Function
ErrHandler:
    fehler = Err.Description

End Function


'===============================================================================
' Name: Sub Delete_Worksheets
' Input:
'   ByRef wb As Workbook
' Output:
'   None
' Purpose: es muss sichergestellt werden, dass immer nur eine Tabelle vorhanden ist
' Remarks: Author H. Sonnenberg Juli 2024
'===============================================================================
Private Sub Delete_Worksheets(ByRef wb As Workbook)
    Dim ctr As Integer
    Dim OG As Integer
    
    On Error Resume Next
    
    If wb.Worksheets.Count > 1 Then
        OG = wb.Worksheets.Count
        Application.DisplayAlerts = False
        For ctr = OG To 1
            wb.Worksheets(ctr).Delete
        Next ctr
        Application.DisplayAlerts = True
    End If
    
End Sub

'===============================================================================
' Name: Function Hole_Template
' Input:
'   ByRef arrTemplates() As String        -   Array mit allen Templates, deren Name, Pfad und Suffix
'   ByVal templateName As String          -   Name des gewünschten Templates
'   ByRef fehler As String                -   Rückgabe eventueller Fehler
' Output:
'   Workbook
' Purpose: anhand eines templateName wird aus dem Array arrTemplates die gewünschte Templatedatei
'       gesucht und wenn gefunden, geöffnet aund als Returnvalue an den Caller übergeben
' Remarks: Author H. Sonnenberg Juli 2024
'===============================================================================
Private Function Hole_Template(ByRef arrTemplates() As String, ByVal templateName As String, ByRef fehler As String) As Workbook
    Dim ctr As Integer
    Dim wb As Workbook
    Dim blnSuccess As Boolean
    
    On Error GoTo ErrHandler
    
    For ctr = LBound(arrTemplates(), 1) To UBound(arrTemplates(), 1)
        If LCase(arrTemplates(ctr, 0)) = LCase(templateName) Then
            templateName = arrTemplates(ctr, 1) & arrTemplates(ctr, 0) & "." & arrTemplates(ctr, 2)
            Set wb = Workbooks.Open(templateName)
            blnSuccess = True
            Exit For
        End If
    Next ctr
    
    If blnSuccess Then
        Set Hole_Template = wb
    Else
        fehler = "Template """ & templateName & """ konnte nicht gefunden werden!"
    End If
Ende:
    Set wb = Nothing
    Exit Function
ErrHandler:
    fehler = Err.Description
    Resume Ende
End Function

'===============================================================================
' Name: Function Create_Kunden_CSV
' Input:
'   ByRef fileName as string
'   ByRef ws As Worksheet
'   ByRef fehler as string
' Output:
'   true / false
' Purpose: CSV Datei für Kunden erzeugen
' Remarks: Author H. Sonnenberg September 2025
'===============================================================================
Function Create_Kunden_CSV(fileName As String, ws As Worksheet, fehler As String) As Boolean
    'fileName:Dir und filename
    'ws: aus welchem Worksheet stammen die Daten
    Dim fileNumber
    Dim sTxt As String
    Dim row As Integer
    Dim strZeit As String
    
    On Error GoTo ErrHandler

    fileNumber = FreeFile

    Open fileName For Output As #fileNumber 'csv öffnen
    On Error GoTo 0
    
    If LCase(ws.Name) = "shortterm_forecast_15min_csv" Then
        sTxt = "TimeStampISO;ModelRunTimeISO;Price_ShortTerm_Forecast_15min_EUR_MWh;"
    ElseIf LCase(ws.Name) = "shortterm_pfc_15min_csv" Then
        sTxt = "TimeStampISO;ModelRunTimeISO;Price_ShortTerm_PFC_15min_EUR_MWh;"
    End If
    Print #fileNumber, sTxt
    
    For row = 2 To 5000
        If ws.Cells(row, 1).Text = vbNullString Then Exit For
        sTxt = ws.Cells(row, 1).Text & ";" & ws.Cells(row, 2).Text & ";" & ws.Cells(row, 3).Text & ";"
        Print #fileNumber, sTxt
    Next row
    
    Create_Kunden_CSV = True

Exit_Function:
    On Error GoTo 0
    On Error Resume Next
    Close #fileNumber 'csv schließen

    Exit Function
ErrHandler:
    fehler = "CSV Datei erzeugen: " & Err.Description
    Resume Exit_Function
End Function

'===============================================================================
' Name: Function Create_Syneco_CSV
' Input:
'   ByRef fileName as string
'   ByRef ws As Worksheet
'   ByRef fehler as string
' Output:
'   true / false
' Purpose: CSV Datei für interne Synecozwecke erzeugen
' Remarks: Author H. Sonnenberg Juni 2024
'===============================================================================
Function Create_Syneco_CSV(fileName As String, ws As Worksheet, fehler As String) As Boolean
    Dim sTxt As String
    Dim i As Integer
    Dim strZeit As String
    Dim fileNumber
    Dim bytErr As Byte: bytErr = 0

    On Error GoTo ErrHandler

    fileNumber = FreeFile

    Open fileName For Output As #fileNumber 'csv öffnen
    On Error GoTo 0

    sTxt = "KurzfristPrognose"
    Print #fileNumber, sTxt
    sTxt = "Zeitachse;Prognosewerte;Base;Peak;OffPeak;"
    Print #fileNumber, sTxt
    For i = 0 To 2200
        If ws.Cells(41 + i, 1).value = vbNullString Then Exit For
        If Len(ws.Cells(41 + i, 1).value) = 10 Then
            strZeit = ws.Cells(41 + i, 1).value & " 00:00:00"
        Else
            strZeit = ws.Cells(41 + i, 1).value
        End If
        On Error Resume Next
        sTxt = strZeit & ";" & ws.Cells(41 + i, 2).value & ";" & ws.Cells(41 + i, 3).value & ";" & ws.Cells(41 + i, 4).value & ";" & ws.Cells(41 + i, 5).value & ";"
        If Err.Number <> 0 Then bytErr = 3: GoTo ErrHandler
        Err.Clear
        On Error GoTo ErrHandler
        Print #fileNumber, sTxt
    Next i
    
    Create_Syneco_CSV = True

Exit_Function:
    On Error GoTo 0
    On Error Resume Next
    Close #fileNumber 'csv schließen

    Exit Function
ErrHandler:
'Resume
    Select Case bytErr
    Case 0:
        fehler = "CSV Datei erzeugen: Es fehlen Preise im Blatt KF_PFC!!!"
    Case 1:
    '    MsgBox "Es fehlt die Quelldatei " & gstrJahr & " !!!", vbCritical + vbOKOnly
    Case 2:
    '    MsgBox "Mappe " & gstrJahr & " ist offen! Bitte schliessen!" & Chr(13) & Chr(13) & "Programm abgebrochen!", vbCritical + vbOKOnly
    Case 3:
        fehler = "CSV Datei erzeugen: Es fehlen Preise im Blatt KF_PFC!!!"
    Case Else
        If Err.Number = 70 Then
            fehler = "CSV Datei erzeugen: " & fileName & " ist noch offen - bitte schliessen!"
        Else
            fehler = "CSV Datei erzeugen: " & Err.Description
        End If
    End Select
    Resume Exit_Function
End Function

'===============================================================================
' Name: Function Get_Pfadangaben
' Input:
'   ByRef einsatzZweck As string
'   ByRef row As Long
'   ByRef pfadAngabe As string
'   ByRef fehler As string
' Output:
'   TRUE/FALSE
' Purpose:  je nach Namensbestandteil des Tool ENT, Test oder keines von
'           beidem werden andere Pfade im Argument pfad zurückgegeben
' Remarks:  Februar 2026 H. Sonnenberg im Rahmen eines Redesign
'===============================================================================
Public Function Get_Pfadangaben(einsatzZweck As String, row As Long, pfadAngabe As String, fehler As String) As Boolean
    Dim rangeName As String
    
    If LCase(einsatzZweck) = "templates" Then
        rangeName = "rngStartTemplates"
    ElseIf LCase(einsatzZweck) = "outputsyneco" Then
        rangeName = "rngStartOutputSyneco"
    ElseIf LCase(einsatzZweck) = "outputkunden" Then
        rangeName = "rngStartOutputKunden"
    End If
    
    On Error GoTo ErrHandler
    
        'Entwicklerpfade in Spalte Q
        If InStr(LCase(ThisWorkbook.Name), "entw") > 0 Then
            pfadAngabe = _
                wsStaticData.Cells(wsStaticData.Range(rangeName).Offset(row, 0).row, _
                wsStaticData.Range(rangeName).Offset(row, 15).Column).value
        'Testpfade in Spalte J
        ElseIf InStr(LCase(ThisWorkbook.Name), "test") > 0 Then
            pfadAngabe = _
                wsStaticData.Cells(wsStaticData.Range(rangeName).Offset(row, 0).row, _
                wsStaticData.Range(rangeName).Offset(row, 8).Column).value
        'produktive Pfade in Spalte D
        ElseIf InStr(LCase(ThisWorkbook.Name), "entw") = 0 And InStr(LCase(ThisWorkbook.Name), "test") = 0 Then
            pfadAngabe = _
                wsStaticData.Cells(wsStaticData.Range(rangeName).Offset(row, 0).row, _
                wsStaticData.Range(rangeName).Offset(row, 2).Column).value
        End If
   
'
'   If LCase(einsatzZweck) = "templates" Then
'        'Entwicklerpfade in Spalte Q
'        If InStr(LCase(ThisWorkbook.Name), "entw") > 0 Then
'            pfadAngabe = _
'                wsStaticData.Cells(wsStaticData.Range("rngStartTemplates").Offset(row, 0).row, _
'                wsStaticData.Range("rngStartTemplates").Offset(row, 15).Column).value
'        'Testpfade in Spalte J
'        ElseIf InStr(LCase(ThisWorkbook.Name), "test") > 0 Then
'            pfadAngabe = _
'                wsStaticData.Cells(wsStaticData.Range("rngStartTemplates").Offset(row, 0).row, _
'                wsStaticData.Range("rngStartTemplates").Offset(row, 8).Column).value
'        'produktive Pfade in Spalte D
'        ElseIf InStr(LCase(ThisWorkbook.Name), "entw") = 0 And InStr(LCase(ThisWorkbook.Name), "test") = 0 Then
'            pfadAngabe = _
'                wsStaticData.Cells(wsStaticData.Range("rngStartTemplates").Offset(row, 0).row, _
'                wsStaticData.Range("rngStartTemplates").Offset(row, 2).Column).value
'        End If
'
'    ElseIf LCase(einsatzZweck) = "outputsyneco" Then
'        'Entwicklerpfade in Spalte Q
'        If InStr(LCase(ThisWorkbook.Name), "entw") > 0 Then
'            pfadAngabe = _
'                wsStaticData.Cells(wsStaticData.Range("rngStartOutputSyneco").Offset(row, 0).row, _
'                wsStaticData.Range("rngStartOutputSyneco").Offset(row, 15).Column).value
'        'Testpfade in Spalte J
'        ElseIf InStr(LCase(ThisWorkbook.Name), "test") > 0 Then
'            pfadAngabe = _
'                wsStaticData.Cells(wsStaticData.Range("rngStartOutputSyneco").Offset(row, 0).row, _
'                wsStaticData.Range("rngStartOutputSyneco").Offset(row, 8).Column).value
'        'produktive Pfade in Spalte D
'        ElseIf InStr(LCase(ThisWorkbook.Name), "entw") = 0 And InStr(LCase(ThisWorkbook.Name), "test") = 0 Then
'            pfadAngabe = _
'                wsStaticData.Cells(wsStaticData.Range("rngStartOutputSyneco").Offset(row, 0).row, _
'                wsStaticData.Range("rngStartOutputSyneco").Offset(row, 2).Column).value
'        End If
'
'    ElseIf LCase(einsatzZweck) = "outputkunden" Then
'
'        'Entwicklerpfade in Spalte Q
'        If InStr(LCase(ThisWorkbook.Name), "entw") > 0 Then
'            pfadAngabe = _
'                wsStaticData.Cells(wsStaticData.Range("rngStartOutputSyneco").Offset(row, 0).row, _
'                wsStaticData.Range("rngStartOutputSyneco").Offset(row, 15).Column).value
'        'Testpfade in Spalte J
'        ElseIf InStr(LCase(ThisWorkbook.Name), "test") > 0 Then
'            pfadAngabe = _
'                wsStaticData.Cells(wsStaticData.Range("rngStartOutputSyneco").Offset(row, 0).row, _
'                wsStaticData.Range("rngStartOutputSyneco").Offset(row, 8).Column).value
'        'produktive Pfade in Spalte D
'        ElseIf InStr(LCase(ThisWorkbook.Name), "entw") = 0 And InStr(LCase(ThisWorkbook.Name), "test") = 0 Then
'            pfadAngabe = _
'                wsStaticData.Cells(wsStaticData.Range("rngStartOutputSyneco").Offset(row, 0).row, _
'                wsStaticData.Range("rngStartOutputSyneco").Offset(row, 2).Column).value
'        End If
'
'    End If

    Get_Pfadangaben = True
    Exit Function
ErrHandler:
    fehler = Err.Description
End Function



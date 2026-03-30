Attribute VB_Name = "modEmail"
'===============================================================================
' Name: modEmail
' Purpose: Modul beinhaltet allgemeine Funktionen und Prozeduren für einen Emailversand
' Remarks: Author H. Sonnenberg im Januar/Februar 2025
' Functions: Hole_Anrede, Hole_Bodytext, Hole_Empfaenger, Hole_LogoDateiFuerMail,
'            Hole_MailRisiko, Hole_PrognoseDatei_FuerKunden, Hole_Sender,
'            Hole_Subject, Hole_URL_Fuer_MailLogo, Hole_User_Infos,
'            Main_Create_Email_For_Customer,
' Properties:
' Methods: Hole_ChartCopy
' Started: 09.01.2026
' Modified:23.02.2026 Umstellung auf Email Versand mit Rest-API MSGRaph
'===============================================================================
Option Explicit

Public Enum enumPriority
    enLow = -1
    enNormal = 0
    enHigh = 1
End Enum

Public Enum TypeOfSender
    enSenderNormal = -1     'User sendet in eigenem Namen, also weder on "Send on Behalf" noch "Send As"
    enSenderOnBehalf = 0    'der Empfäger erkennt anhand der Nachricht, das der Benutzer im Namen eines anderen gesendet hat
    enSenderSendAs = 1      'Es gibt in der fertigen Mail keinen Hinweis darauf, das der Benutzer im Namen eines anderen gesendet hat
End Enum


'===============================================================================
' Name: Function Hole_LogoDatei
' Input:
'   ByRef logoDatei As String
'   ByRef fehler As String
' Output:
'   True / False
' Purpose: liefert Name und Pfad einer möglichen Logodatei, wenn in Staticdata angegeben
' Remarks: Author H. Sonnenberg Januar/Februar 2025
'===============================================================================
Public Function Hole_LogoDateiFuerMail(logoDatei As String, fehler As String) As Boolean
    
    On Error GoTo ErrHandler
    
    If wsStaticData.Range("rngLogoDatei").value <> vbNullString Then _
        logoDatei = wsStaticData.Range("rngLogoDatei").value Else logoDatei = vbNullString
    
    Hole_LogoDateiFuerMail = True
    
    Exit Function
ErrHandler:
    fehler = Err.Description
End Function

'===============================================================================
' Name: Function Hole_URL_Fuer_MailLogo
' Input:
'   ByRef selPfadDatei As String
' Output:
'   True / False
' Purpose: liefert die URL einer möglichen Website, wenn in Staticdata angegeben
' Remarks: Author H. Sonnenberg Januar/Februar 2025
'===============================================================================
Public Function Hole_URL_Fuer_MailLogo(URL As String, fehler As String) As Boolean
    
    On Error GoTo ErrHandler
    
    If wsStaticData.Range("rngLogoURL").value <> vbNullString Then _
        URL = wsStaticData.Range("rngLogoURL").value Else URL = vbNullString
    
    Hole_URL_Fuer_MailLogo = True
    
    Exit Function
ErrHandler:
    fehler = Err.Description
End Function

'===============================================================================
' Name: Function Hole_Sender
' Input:
'   ByRef recipients As String
'   ByRef fehler As String
' Output:
'   True / False
' Purpose: Referenzübergabe: eine mögliche Senderangabe, wenn der aktuelle
'          Benutzer im Auftrag dieses Senders die Mail verschickt,
'          aktuell nicht verwendet
' Remarks: Author H. Sonnenberg Januar/Februar 2025
'===============================================================================
Public Function Hole_Sender(sender() As String, fehler As String) As Boolean
    On Error GoTo ErrHandler
    
    sender() = Split(Trim(wsStaticData.Range("rngSender").value), ";")
    
    
    Hole_Sender = True
    
    Exit Function
ErrHandler:
    fehler = Err.Description
    
End Function




'===============================================================================
' Name: Function Hole_Empfaenger
' Input:
'   ByRef recipients As String
'   ByRef fehler As String
' Output:
'   True / False
' Purpose: liest den/die Empfänger aus Staticdata aus
' Remarks: Author H. Sonnenberg Januar/Februar 2025
'===============================================================================
Public Function Hole_Empfaenger(Recipients() As String, fehler As String) As Boolean
    'hartmutsonnenberg@posteo.de;info@hartmutsonnenberg.com;h.sonnenberg@vba-solutions.de
    'aber auch spot@syneco.net ist aber zugleich Versender der Email
    On Error GoTo ErrHandler
    
    
    
    Recipients() = Split(wsStaticData.Range("rngRecipients").value, ";")

    Hole_Empfaenger = True
    
    Exit Function
ErrHandler:
    fehler = Err.Description
    
End Function


'===============================================================================
' Name: Function Hole_MailRisiko
' Input:
'   ByRef mailRisiko As String
'   ByRef fehler As String
' Output:
'   True / False
' Purpose: falls die Plausibilitätsprüfung Fehler wirft, mit der Veröffentlichung
'          auch eine Mail an Risiko versendet, diese Adresse wird hier per
'          Referenzübergabe zurückgegeben
' Remarks: Author H. Sonnenberg Januar/Februar 2025
'===============================================================================
Public Function Hole_MailRisiko(mailRisiko() As String, fehler As String) As Boolean
    On Error GoTo ErrHandler
    
    mailRisiko() = Split(Trim(wsStaticData.Range("rngMailRisiko").value), ";")
    
    Hole_MailRisiko = True
    
    Exit Function
ErrHandler:
    fehler = Err.Description
End Function


'===============================================================================
' Name: Function Hole_PrognoseDatei_FuerKunden
' Input:
'   ByRef prognoseDatei As String
'   ByRef fehler As String
' Output:
'   True / False
' Purpose:  die während des Programmablaufes erzeugte Datei "Kurzfrist_hPFC_23_25.xlsx"
'           (siehe staticdata rngPrognoseDatei) soll als Anhang der Email mitgegeben werden
' Remarks:  Author H. Sonnenberg Januar/Februar 2025
'===============================================================================
Public Function Hole_PrognoseDatei_FuerKunden(prognoseDatei As String, fehler As String) As Boolean
    On Error GoTo ErrHandler
    
    'der Email mitgegeben werden
    'K:\XTRANET_TRANSFER\Daten\01_Marktueberblick\Marktdaten\Kurzfristprognose\Kurzfrist_hPFC_23_25.xlsx
    
'    If InStr(LCase(ThisWorkbook.Name), "entw") > 0 Then
'        prognoseDatei = "C:\Users\Sonnenberg_H\Desktop\GasTrades.xls"
'    Else
        prognoseDatei = CheckPathName(wsStaticData.Range("rngPrognoseDatei").Offset(0, -1).value) & wsStaticData.Range("rngPrognoseDatei").value & "." & wsStaticData.Range("rngPrognoseDatei").Offset(0, 1).value
        If Dir(prognoseDatei) = "" Then Err.Raise 901, "Hole_PrognoseDatei", "Datei: """ & prognoseDatei & """ nicht gefunden"
'    End If

    Hole_PrognoseDatei_FuerKunden = True
    
    Exit Function
ErrHandler:
    fehler = Err.Description

End Function



'===============================================================================
' Name: Function Hole_Bodytext
' Input:
'   ByRef prognoseDatei As String
'   ByRef fehler As String
' Output:
'   True / False
' Purpose: liest den Bodytext aus Staticdata aus und liefert per Referenzübergabe
' Remarks: Author H. Sonnenberg Januar/Februar 2025
'===============================================================================
Public Function Hole_Bodytext(BodyText As String, fehler As String) As Boolean
    On Error GoTo ErrHandler
    
    BodyText = wsStaticData.Range("rngBodyText").value
    Hole_Bodytext = True
    
    Exit Function
ErrHandler:
    fehler = Err.Description
End Function

'===============================================================================
' Name: Function Hole_Anrede
' Input:
'   ByRef anrede As String
'   ByRef fehler As String
' Output:
'   True / False
' Purpose: holt den Anredetext
' Remarks: Author H. Sonnenberg Januar/Februar 2025
'===============================================================================
Private Function Hole_Anrede(anrede As String, fehler As String) As Boolean
    On Error GoTo ErrHandler
    
    anrede = wsStaticData.Range("rngAnrede").value
    Hole_Anrede = True
    
    Exit Function
ErrHandler:
    fehler = Err.Description
End Function


'===============================================================================
' Name: Function Hole_User_Infos
' Input:
'   ByRef oGraph As clsMSGraph
'   ByRef fehler As String
' Output:
'   True / False
' Purpose:  fragt über MSGraph alle für die Email-Signatur wichtigen Attribute
'           des User Ressourcentyps ab. Das sind displayName, department, companyname,
'           streetaddress, postalCode, city, businessphones, mail. Die Angaben
'           werden innerhalb der Klasse gekapselt im Array userInfos für spätere
'           Verwendung festgehalten.
' Remarks: Author H. Sonnenberg Januar/Februar 2025
'===============================================================================
Private Function Hole_User_Infos(oGraph As clsMSGraph, fehler As String) As Boolean
    Dim fehlerIntern As String
    On Error GoTo ErrHandler
    
    If oGraph.Request_UserInfos(fehlerIntern) = False Then
        fehler = fehlerIntern
    Else
        Hole_User_Infos = True
    End If
    
    Exit Function
ErrHandler:
    fehler = Err.Description

End Function



'===============================================================================
' Name: Function Hole_Subject
' Input:
'   ByRef prognoseDatei As String
'   ByRef fehler As String
' Output:
'   True / False
' Purpose: holt den Betreff
' Remarks: Author H. Sonnenberg Januar/Februar 2025
'===============================================================================
Private Function Hole_Subject(Subject As String, fehler As String) As Boolean
    On Error GoTo ErrHandler
    
    Subject = wsStaticData.Range("rngBetreff").value
    Hole_Subject = True
    
    Exit Function
ErrHandler:
    fehler = Err.Description
End Function



'===============================================================================
' Name: Function Main_Create_Email_For_Customer
' Input:
'   ByRef foreCastFile As String
'   ByRef fehler As String
' Output:
'   True / False
' Purpose:  bereitet die Erzeugung der Email vor, indem alle für die aktuelle
'           Situation nötigen Daten beschafft und mit Hilfe der Properties der
'           Klasse  übergeben werden
'           Der Aufruf der Methoden Email_Create und Email_Send schliesst den
'           Versand ab
' Remarks: Author H. Sonnenberg Januar/Februar 2026
'===============================================================================
Public Function Main_Create_Email_For_Customer(foreCastFile As String, fehler As String) As Boolean
    Dim oGraph As New clsMSGraph
    Dim varEmpfaenger As Variant
    Dim varEmpfaengerKopie As Variant
    Dim varEmpfaengerBlindKopie As Variant
    'wenn beide False sind, sendet A als A
    Dim blnSendOnBehalf As Boolean                              'wird im Auftrag einer anderen Person gesendet? Also Person A im Auftrag von B?
    Dim blnSendAs As Boolean                                    'sendet Person A als B?
    Dim Recipients() As String                                  'Empfänger
    Dim sender() As String                                      'Versender der Email: muss nicht zwingend der angemeldete Benutzer sein
    Dim ctrArray As Integer
    Dim BodyText As String                                      'Mailtext incl. encodierter Images (Diagramm und evtl. Logo)
    Dim Subject As String                                       'Betreff
    Dim encodedFile As String                                   'Base 64 encodierte Datei
    Dim blnSuccess As Boolean                                   'Base 64 Encodierung erfolgreich?
    Dim anrede As String                                        'Anrede Email
    Dim mailRisiko() As String                                  'Mailadresse der Abteilung Risiko
    Dim logoDatei As String                                     'Pfad und Name zu einer eventuellen Logodatei
    Dim LogoURL As String                                       'URL, die mit Mausklick Logo aktiviert wird
    Dim fehlerCreateEmail As String

    On Error GoTo ErrHandler
    
    blnSendAs = False 'True
    blnSendOnBehalf = False
    
    '1. Vor allem anderen: erstmal versuchen, ob der Authorisierungscode fehlerfrei übermittelt wird
    With oGraph
        .ClientID = wsStaticData.Range("rngClientID").value
        If blnSendOnBehalf = False And blnSendAs = False Then
            .Permissions.Add "mail.send"
        ElseIf blnSendOnBehalf = True Or blnSendAs = True Then
            .Permissions.Add "mail.send.Shared"
        End If
        .useRefreshToken = True
        .LinkPermissions '       alle permissions zu einem Scope vereinigen
         
        If .Request_Authorization(fehlerCreateEmail) = False Then Err.Raise 901, "Autorisierung", fehlerCreateEmail
    End With
    
    '2. Token anfordern
    With oGraph
        If .Request_Token(fehlerCreateEmail) = False Then Err.Raise 901, "request Token", fehlerCreateEmail
    End With
    
    '3. Daten einsammeln und versenden
    If Hole_User_Infos(oGraph, fehlerCreateEmail) = False Then Err.Raise 902, "send Email", fehlerCreateEmail
    If Hole_Empfaenger(Recipients, fehlerCreateEmail) = False Then Err.Raise 902, "send Email", fehlerCreateEmail
    If Hole_Bodytext(BodyText, fehlerCreateEmail) = False Then Err.Raise 902, "send Email", fehlerCreateEmail
    If Hole_Subject(Subject, fehlerCreateEmail) = False Then Err.Raise 902, "send Email", fehlerCreateEmail
    If Hole_Anrede(anrede, fehlerCreateEmail) = False Then Err.Raise 902, "send Email", fehlerCreateEmail
    If Hole_Sender(sender, fehlerCreateEmail) = False Then Err.Raise 902, "send Email", fehlerCreateEmail
    If Hole_MailRisiko(mailRisiko, fehlerCreateEmail) = False Then Err.Raise 902, "send Email", fehlerCreateEmail
    If Hole_LogoDateiFuerMail(logoDatei, fehler) = False Then Err.Raise 902, "send Email", fehlerCreateEmail
    If Hole_URL_Fuer_MailLogo(LogoURL, fehler) = False Then Err.Raise 902, "send Email", fehlerCreateEmail
    
    With oGraph
        
        If blnSendOnBehalf = False And blnSendAs = False Then
            .SenderType = enSenderNormal
        ElseIf blnSendOnBehalf = False And blnSendAs = True Then
            .SenderType = enSenderSendAs
            For ctrArray = 0 To UBound(sender())
                .SenderAs.Add Trim(sender(ctrArray))
            Next ctrArray
        ElseIf blnSendOnBehalf = True And blnSendAs = False Then
            .SenderType = enSenderOnBehalf
            .SenderOnBehalf.Add sender
        End If
        
     '   .UserEmailOnBehalf = sender                         'Versender
        For ctrArray = 0 To UBound(Recipients())
            .Recipients.Add Trim(Recipients(ctrArray))      'gegen unwillkommene Leerzeichen
        Next ctrArray
        
        
        For ctrArray = 0 To UBound(sender())
            .CcRecipients.Add Trim(sender(ctrArray))      'gegen unwillkommene Leerzeichen
        Next ctrArray
        
        
        'wenn die Plausibilitätsprüfung nicht Nullstring ergibt, dann ebenfalls eine Mail an Risiko in cc.
        If CheckMeldungen = False Then
            For ctrArray = 0 To UBound(mailRisiko())
                .CcRecipients.Add Trim(mailRisiko(ctrArray))
            Next ctrArray
        End If
        
        .Subject = "Betreff: " & Subject
        
        'kopiert das Chart-Objekt im Sheet "KF_PFC_W4_23_25" in den Bodytext hinein
        'erzeugt ein encodiertes Image eines Diagramms
        encodedFile = Create_Encoded_ChartFile("KF_PFC_W4_23_25", "Diagramm 3", blnSuccess)
        If blnSuccess = False Then GoTo Ende
        'hier wieder zurückstellen
        oGraph.EmbeddedImage = True
        oGraph.encodedFile = encodedFile
       
        .BodyText = anrede & "</br></br>" & BodyText & vbCrLf & vbCrLf
        
        .ShowImpressum = True
        
        If LCase(Trim(wsStaticData.Range("rngSignatur").value)) = "ja" Then .ShowSignatur = True Else .ShowSignatur = False
        
        If logoDatei <> vbNullString Then
            If Dir(logoDatei, vbNormal) <> vbNullString Then
                encodedFile = Create_Encoded_Picture_File(logoDatei, blnSuccess)
                If blnSuccess = False Then GoTo Ende
                
                oGraph.EmbeddedLogo = True
                oGraph.encodedLogo = encodedFile
            End If
            If LogoURL <> vbNullString Then
                oGraph.LogoURL = LogoURL
            End If
        End If
        
        .Attachments.Add foreCastFile
        
        .Priority = enNormal
        
        '"HTML Email erzeugen: Bodytext muss in 2 Teile mit Chart dazwischen entwickelt werden!
        If .Email_Create(fehlerCreateEmail) = False Then Err.Raise 901, "HTML Email erzeugen", fehlerCreateEmail
        ' dann Email_Send
        If .Email_Send(fehler:=fehlerCreateEmail) = False Then Err.Raise 901, "HTML Email versenden", fehlerCreateEmail
    
    End With
    
    Main_Create_Email_For_Customer = True
    
Ende:
    Set oGraph = Nothing
    
    Exit Function
    
ErrHandler:
    If fehlerCreateEmail <> "" Then
        fehler = fehlerCreateEmail
    ElseIf Err.Number <> 0 Then
        fehler = Err.Description
    End If
    
    Resume Ende
    
End Function


'===============================================================================
' Name: Sub Hole_ChartCopy
' Input:
'   ByRef ws As Worksheet
' Output:
'   None
' Purpose: kopiert das Diagramm aus der angegebenen Tabelle in die Zwischenablage
' Remarks: Author H. Sonnenberg Januar/Februar 2025
'===============================================================================
Private Sub Hole_ChartCopy(ws As Worksheet)
    ws.Select
    ws.Unprotect gstrPWD
    
    ws.Range("G35").Select
    ws.ChartObjects("Diagramm 3").Activate
    ActiveChart.ChartArea.Copy

    ws.Protect gstrPWD
    
End Sub



'===============================================================================
' Name: Function Create_Encoded_ChartFile
' Input:
'   ByRef wsName As String
'   ByRef chartName As String
'   ByRef success As Boolean
' Output:
'   String
' Purpose: ein Chartobjekt wird Base64 encodiert und als File gespeichert
' Remarks: Author H. Sonnenberg Januar/Februar 2026
'===============================================================================
Private Function Create_Encoded_ChartFile(wsName As String, chartName As String, success As Boolean) As String
    Dim o_Image As clsEncodeImage
    Dim fehler As String
    Dim savedFile As String                                     '
    Dim encodedFile As String                                     '

    
    On Error GoTo ErrHandler
    'kopiert das Chart-Objekt im Sheet "KF_PFC_W4_23_25" in den Bodytext hinein
    Set o_Image = New clsEncodeImage

    o_Image.nameWorkSheet = wsName '"KF_PFC_W4_23_25"
    o_Image.chartName = chartName '"Diagramm 3"

    If Dir(ThisWorkbook.path & "\images", vbDirectory) = "" Then MkDir ThisWorkbook.path & "\images"
    o_Image.DirectoryToSave = ThisWorkbook.path & "\images"

    If o_Image.SaveChartToFile(savedFile, fehler) = False Then Err.Raise 901, "Chart speichern", fehler

    If o_Image.EncodeFile(savedFile, encodedFile, fehler) = False Then Err.Raise 901, "Chart speichern", fehler
    
    success = True
    
    Create_Encoded_ChartFile = encodedFile

Ende:
    Set o_Image = Nothing
    Exit Function
ErrHandler:
    MsgBox chartName & " in Tabelle """ & wsName & """ konnte nicht in die Mail aufgenommen werden!" / vbCrLf & vbCrLf & Err.Description, vbCritical, ThisWorkbook.Name
    Resume Ende
End Function



'===============================================================================
' Name: Function Create_Encoded_Picture_File
' Input:
'   ByRef picture As String
'   ByRef success As Boolean
' Output:
'   String
' Purpose: eine Bilddatei (hier Logo) wird Base64 encodiert und als File gespeichert
' Remarks: Author H. Sonnenberg Januar/Februar 2026
'===============================================================================
Private Function Create_Encoded_Picture_File(Picture As String, success As Boolean) As String
    Dim o_Image As clsEncodeImage
    Dim fehler As String
    Dim savedFile As String                                     '
    Dim encodedFile As String                                     '
    
    On Error GoTo ErrHandler
    
    Set o_Image = New clsEncodeImage

    If Dir(ThisWorkbook.path & "\images", vbDirectory) = "" Then MkDir ThisWorkbook.path & "\images"
    o_Image.DirectoryToSave = ThisWorkbook.path & "\images"
    
    'Bilddatei umwandeln und abspeichern
    If o_Image.EncodeFile(Picture, encodedFile, fehler) = False Then Err.Raise 901, "Chart speichern", fehler
    
    success = True
    
    Create_Encoded_Picture_File = encodedFile

Ende:
    Set o_Image = Nothing
    Exit Function
ErrHandler:
    MsgBox "Bilddatei """ & Picture & """ konnte nicht in die Mail aufgenommen werden!" / vbCrLf & vbCrLf & Err.Description, vbCritical, ThisWorkbook.Name
    Resume Ende
End Function



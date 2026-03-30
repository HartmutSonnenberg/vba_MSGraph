VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMSGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name: clsMSGraph
' Purpose:      Klasse beinhaltet
' Remarks:      Author H. Sonnenberg im Januar 2026
' Properties:   read:           Attachments, Recipients, BccRecipients, CcRecipients,
'                               Authorization_Code, ExpiresIn, Permissions,
'                               RefreshToken, SenderAs, SenderOnBehalf, Token
' Properties:   write:          EmbeddedImage, encodedFile, SenderType, Priority,
'                               SenderIsAnotherUser, useRefreshToken, UserEmailOnBehalf
'                               ShowSignatur, ShowImpressum
' Properties:   read, write:    BodyText, Subject
' Properties:   locked:         m_DisplayName, m_Department , m_CompanyName, m_Streetaddress,
'                               m_PostalCode, m_City, m_Businessphones, m_Mobilephone, m_Mail
' Methods:      public:         Email_Create, Email_Send, EncodeFileToBase64,
'                               LinkPermissions, Request_Token, Request_Authorization
'                               Request_UserInfos
' Methods:      private:        Get_AuthorizationCode, GetFilename
' Started:      27.12.2025
' Modified:     12.11.2026      Kommentare eingefügt, vereinfacht
'===============================================================================
'send on behalf oder sand as:
'https://learn.microsoft.com/de-de/graph/outlook-send-mail-from-other-user

''https://learn.microsoft.com/de-de/graph/outlook-send-mail-from-other-user
'Wenn Ihre Anwendung die Nachricht über den /me-Endpunkt sendet (oder über /users/{user-id}, wobei user-id dem angemeldeten Benutzer entspricht),
'so wird die Nachricht standardmäßig im Ordner „Gesendete Elemente“ des sendenden Benutzers gespeichert.
'Wenn Ihre Anwendung die Nachricht über /users/{user-id} sendet, wobei user-id dem Benutzer entspricht, von dem die Nachricht stammt,
'so wird die Nachricht standardmäßig im Ordner „Gesendete Elemente“ des Benutzers gespeichert, von dem die Nachricht stammt.

''https://learn.microsoft.com/de-de/graph/api/resources/mail-api-overview?view=graph-rest-1.0

''https://learn.microsoft.com/en-us/graph/outlook-mail-concept-overview

''https://learn.microsoft.com/en-us/graph/permissions-reference
'Mail.Send.Shared is only valid for work or school accounts.


'toRecipients,


Option Explicit


Private m_UserEmailOnBehalf As String           'die Emailadresse des aktuellen Users, wird meist nicht verwendet, da wir i.d.R. den me-Endpunkt verwenden und
                                                'der impliziert einen angemeldeten Benutzer

Private m_blnSenderIsAnotherUser As Boolean     'Steuervariable für den Fall einer Nachricht im Namen oder als ein anderer Nutzer/Gruppe
                                                'Property SenderIsAnotherUser wird aktuell nicht verwendet

Private m_SenderType As TypeOfSender            'denkbar sind drei Typen Normal, OnBehalf und SendAs - siehe Enum TypeOfSender

Private m_Response_Type As String               'muss für den OAuth 2.0-Autorisierungscodefluss den Wert "code" haben. In anderen Zusammenhängen kann es "token" sein

Private m_Response_Mode As String               'Methode der Rückgabe an die App: query oder form_Post

Private m_Scope As String                       'Eine durch Leerzeichen getrennte Liste von Microsoft Graph-Berechtigungen, denen der Benutzer zustimmen soll.
                                                'Diese Berechtigungen können Ressourcenberechtigungen wie User.Read und Mail.Read sowie OIDC-Bereiche wie
                                                'offline_access enthalten, was angibt, dass die App ein Aktualisierungstoken für den langlebigen Zugriff auf Ressourcen benötigt
            
Private m_state As String                       'Ein in der Anforderung enthaltener Wert, der auch in der Tokenantwort zurückgegeben wird. Dabei kann es sich um eine Zeichenfolge
                                                'mit beliebigen Inhalten handeln. Ein zufällig generierter eindeutiger Wert wird in der Regel verwendet , um siteübergreifende
                                                'Anforderungsfälschungsangriffe zu verhindern.

Private m_Permissions As Collection             'enthält eine Liste aller gewünschten Berechtigungen wie mail.read, mail.send etc. !

Private m_Authorization_Code As String          'enthält den Autorisierungscode im Code-Parameter, wird in der Zeichenfolge der Umleitungs-URL zurückgegeben, read-only!

Private m_Token As String                       'enthält den Zugriffs-Token

Private m_Use_RefreshToken As Boolean           'soll Refresh_Token angefordert werden

Private m_RefreshToken As String                'enthält den Refresh_Token

Private m_ExpiresInValue As Integer             'Gültigkeitsdauer des Token in Sekunden (üblicherweise zwischen 60 und 90 Minuten - 3600 bis 5400 Sekunden)

Private m_Subject As String                     'der Betreff

Private m_Priority As enumPriority              'es gibt die Prioritäten "low", "normal" oder "high"

Private m_SenderOnBehalf As Collection          'Versender on Behalf der Email

Private m_SenderAs As Collection                'Sender als Versender der Email

Private m_Recipients As Collection              'alle Empfänger der Email

Private m_CcRecipients As Collection            'Kopie an ( = CC)

Private m_BccRecipients As Collection           'Blindkopie (BCC)

Private m_Attachments As Collection             'Dateianhäge

Private m_BodyText As String                    'Bodytext wird von der Funktion Email_Create erzeugt

Private m_encodedFile As String                 'wenn im Bodytext ein Image mitgeschickt werden soll, das Base64 umgewandelte Image

Private m_blnEmbeddedImage As Boolean           'wenn im Bodytext ein Image mitgeschickt werden soll

Private m_encodedLogo As String                 'wenn am Ende der Mail ein Logo mitgeschickt werden soll, das Base64 umgewandelte Logo

Private m_blnEmbeddedLogo As Boolean            'wenn am Ende der Mail ein Logo mitgeschickt werden soll

Private m_strlogoURL As String                  'wenn auf dem Logo ein Link zu einer Website untergebracht werden soll

Private m_JSON_Email As String                  'die komplette Email als JSON Text, bereit für den Versand (siehe Methode Email_Create!)

Private m_ClientID As String                    'die Client-ID der Syneco MS Graph App in MS Entra

Private m_blnShowImpressum As Boolean           'soll in der Mail ein Impressum gesendet werden

Private m_blnShowSignatur As Boolean            'soll in der Mail eine Signatur gesendet werden

Private m_DisplayName As String                 'Benutzername und Vorname des user Ressourcentyps: für Email

Private m_Department As String                  'department-Attribut des user Ressourcentyps: für Email

Private m_Companyname As String                 'companyname-Attribut des user Ressourcentyps: für Email

Private m_Streetaddress As String               'streetaddress-Attribut des user Ressourcentyps: für Email

Private m_PostalCode As String                  'PLZ-Attribut des user Ressourcentyps: für Email

Private m_City As String                        'Ort-Attribut des user Ressourcentyps: für Email

Private m_Businessphones As String              'Festnetz-Attribut des user Ressourcentyps: für Email

Private m_Mobilephone As String                 'Mobilfunk-Attribut des user Ressourcentyps: für Email

Private m_Mail As String                        'Mailadresse-Attribut des user Ressourcentyps: für Email

Private Const cTenant_ID = _
    "0c8e694b-ca0b-4caf-81a7-fd7ed0aac68e"      'Syneco Tenant-ID

Private Const cRedirect_URL = _
    "https://www.testoutlook.de"                'ist in der Syneco MS Graph App vorgegeben, nur durch Tenant-Admin änderbar

Private Const c_Blank = "%20"                   '



'******************************************************************************************************************************************
'************************************************        Eigenschaften       **************************************************************
'******************************************************************************************************************************************
'write-only
Public Property Let ClientID(appClientID As String)
    m_ClientID = appClientID
End Property
'write-only
Public Property Let useRefreshToken(useRefreshToken As Boolean)
    m_Use_RefreshToken = useRefreshToken
End Property
'write-only
Public Property Let UserEmailOnBehalf(aktUserEmailOnBehalf As String)
    m_UserEmailOnBehalf = aktUserEmailOnBehalf
End Property
'write-only
Public Property Let encodedFile(strEncodedFile As String)
    m_encodedFile = strEncodedFile
End Property
'write-only
Public Property Let EmbeddedImage(blnEmbeddedImage As Boolean)
    m_blnEmbeddedImage = blnEmbeddedImage
End Property
'write-only
Public Property Let ShowImpressum(blnShowImpressum As Boolean)
    m_blnShowImpressum = blnShowImpressum
End Property
'write-only
Public Property Let ShowSignatur(blnShowSignatur As Boolean)
    m_blnShowSignatur = blnShowSignatur
End Property
'write-only
Public Property Let encodedLogo(strLogoDatei As String)
    m_encodedLogo = strLogoDatei
End Property
'write-only
Public Property Let EmbeddedLogo(blnEmbeddedLogo As Boolean)
    m_blnEmbeddedLogo = blnEmbeddedLogo
End Property
'write-only
Public Property Let LogoURL(strLogoURL As String)
    m_strlogoURL = strLogoURL
End Property

'alle Properties vom Typ Collection können anders behandelt werden, wenn neue Elemente angefügt werden sollen
'es genügt die Property Get, denn habe ich eine Referenz auf das Objekt, kann ich dessen Add-Methode aufrufen
Public Property Get Recipients() As Collection
    Set Recipients = m_Recipients
End Property

Public Property Get SenderOnBehalf() As Collection
    Set SenderOnBehalf = m_SenderOnBehalf
End Property

Public Property Get SenderAs() As Collection
    Set SenderAs = m_SenderAs
End Property

Public Property Get Permissions() As Collection
    Set Permissions = m_Permissions
End Property

Public Property Get CcRecipients() As Collection
    If m_CcRecipients Is Nothing Then Set m_CcRecipients = New Collection
    Set CcRecipients = m_CcRecipients
End Property

Public Property Get BccRecipients() As Collection
    If m_BccRecipients Is Nothing Then Set m_BccRecipients = New Collection
    Set BccRecipients = m_BccRecipients
End Property

Public Property Get Attachments() As Collection
    If m_Attachments Is Nothing Then Set m_Attachments = New Collection
    Set Attachments = m_Attachments
End Property
'Ende der Properties vom Typ Collection
Public Property Let Priority(varPriority As enumPriority)
    m_Priority = varPriority
End Property
'write-only
Public Property Let SenderIsAnotherUser(blnSenderIsAnotherUser As Boolean)
    m_blnSenderIsAnotherUser = blnSenderIsAnotherUser
End Property
'write-only
Public Property Let SenderType(par_senderType As TypeOfSender)
    m_SenderType = par_senderType
End Property
'write-only
Public Property Let Subject(varSubject As String)
    m_Subject = varSubject
End Property
'write-only
Public Property Let BodyText(varBodyText As String)
    m_BodyText = varBodyText
End Property
'read-only
Public Property Get Authorization_Code() As String
    Authorization_Code = m_Authorization_Code
End Property
'read/write
Public Property Get Token() As String
    Token = m_Token
End Property
'read/write
Public Property Let Token(aToken As String)
    m_Token = aToken
End Property
'read-only
Public Property Get RefreshToken() As String
    RefreshToken = m_RefreshToken
End Property
'read-only
Public Property Get ExpiresIn() As Integer
    ExpiresIn = m_ExpiresInValue
End Property


'******************************************************************************************************************************************
'************************************************          Methoden          **************************************************************
'******************************************************************************************************************************************
'===============================================================================
' Name: Function Email_Create
' Input:
'   ByRef fehler As String      - Verwendung by Ref!
' Output:
'   True/False
' Purpose:  Es werden alle internen Collections mit Elementen verwendet, diese in
'           Dictionaries einzeln übertragen, diese in Collections eingehängt und
'           abschließend an ein übergeordnetes Message Objekt übergeben
'           Anschließend mit ConvertToJson in einen JSON String umgewandelt
'           C#: aus Collection wird List(<T> und aus VBA-Dictionary wird ein
'               Dictionary<TKey, TValue>
' Remarks: Author H. Sonnenberg Januar 2026
'===============================================================================
Public Function Email_Create(fehler As String) As Boolean
    'die komplette Email als JSON String
    Dim dicMail As New Dictionary
    Dim dicMessage As New Dictionary
    Dim dicBody As New Dictionary

    'Empfänger
    Dim colToRecipients As New Collection               'Collection aller Objekte Empfänger
    Dim dicToRecipient As Dictionary                    'Empfänger Dictionary
    Dim dicToEMailAddress As Dictionary                 'Dictionary der Empfänger Emailadressen
    
    'Empfänger Kopie
    Dim colToCCRecipients As Collection                 'Collection aller Objekte CC-Empfänger
    Dim dicToCCRecipient As Dictionary                  'CC_Empfänger Dictionary
    Dim dicToCCEMailAddress As Dictionary               'Dictionary der CC_Empfänger Emailadressen
    
    'Empfänger Blind-Kopie
    Dim colToBCCRecipients As Collection                'Collection aller Objekte BCC-Empfänger
    Dim dicToBCCRecipient As Dictionary                 'BCC_Empfänger Dictionary
    Dim dicToBCCEMailAddress As Dictionary              'Dictionary der BCC_Empfänger Emailadressen
    
    'Versender on behalf
    Dim colOnBehalf As Collection                       'Collection aller Objekte Versender (on behalf)
    Dim dicSenderOnBehalf As Dictionary                 'Dictionary der Versender (on behalf)
    Dim dicToAddressOnBehalf As Dictionary              'Dictionary der Versender (Send As) Emailadressen
    
    'Versender Send As
    Dim colSenderAs As Collection                       'Collection aller Objekte Versender (on behalf)
    Dim dicSenderAs As Dictionary
    Dim dicToAddressSenderAs As Dictionary              'Dictionary der Versender (on behalf) Emailadressen

    'Dateianhänge
    Dim colAttachments As Collection                    'Collection aller Attachments
    Dim dicAttachments As Dictionary                    'Dictionary der Attachments
    
    Dim lng As Integer
    Dim BodyText As String
    Dim strFileContent As String
    Dim strJSON As String
    
    Dim impressum As String                             'hier sind Signatur und Impressum in einem String gemeint
    
    On Error GoTo ErrHandler
    
    dicMail.Add "message", dicMessage
    dicMessage.Add "subject", m_Subject
    dicMessage.Add "body", dicBody
    dicBody.Add "contentType", "HTML" '"text" im reinen Text-Modus kann die Grafik nicht im Bodytext mitgeschickt werden!
    
    
    If m_blnEmbeddedImage = True Then
        m_BodyText = m_BodyText & _
            "<p><img src=""data:image/png;base64," & m_encodedFile & " alt=""Chart"" /></p>"
    End If
    
    m_BodyText = m_BodyText & vbCrLf & vbCrLf '& "----------------------------------------------------------------------------------------------------" & vbCrLf
    
    
    If m_blnShowSignatur And m_blnShowImpressum Then
        impressum = Mail_Add_In.Get_ImpressumAndSignature
        
        'im Impressums- und Signaturtext die Platzhalter durch die Inhalte der UserInfo Properties ersetzen
        impressum = Replace(impressum, "Name MA", m_DisplayName)
        impressum = Replace(impressum, "Abteilung", m_Department)
        impressum = Replace(impressum, "Strasse", m_Streetaddress)
        impressum = Replace(impressum, "PLZ", m_PostalCode)
        impressum = Replace(impressum, "Ort", m_City)
        impressum = Replace(impressum, "Festnetz_MA", m_Businessphones)
        impressum = Replace(impressum, "Mobil_MA", m_Mobilephone)
        impressum = Replace(impressum, "Email_MA", m_Mail)
        impressum = Replace(impressum, "webAdresse_Syneco", "www.syneco.net")

    ElseIf m_blnShowSignatur = False And m_blnShowImpressum Then
        impressum = Mail_Add_In.Get_Impressum
    End If
    
    If m_blnShowImpressum Or m_blnShowSignatur Then
        If InStr(m_BodyText, impressum) = 0 Then m_BodyText = m_BodyText & impressum
    End If
    
    If m_blnEmbeddedLogo And m_strlogoURL = vbNullString Then
        If InStr(m_BodyText, "data:image/png;base64") = 0 Then _
            m_BodyText = m_BodyText & "</br>" & _
                "<p><img src=""data:image/png;base64," & m_encodedLogo & " alt=""Logo"" /></p>"
    ElseIf m_blnEmbeddedLogo And m_strlogoURL <> vbNullString Then
        If InStr(m_BodyText, "data:image/png;base64") = 0 Then _
        m_BodyText = m_BodyText & "</br>" & _
            "<p><a href=" & m_strlogoURL & "><img src=""data:image/png;base64," & m_encodedLogo & " alt=""Logo"" /></a></p>"
    End If
    
    dicBody.Add "content", m_BodyText
    
    If m_SenderType = enSenderOnBehalf Then
        'wer sendet?
        Set colOnBehalf = New Collection
        dicMessage.Add "sender", colOnBehalf
        For lng = 1 To m_SenderOnBehalf.Count
            Set dicSenderOnBehalf = New Dictionary
            colOnBehalf.Add dicSenderOnBehalf
            Set dicToAddressOnBehalf = New Dictionary
            dicSenderOnBehalf.Add "emailAddress", dicToAddressOnBehalf
            dicToAddressOnBehalf.Add "name", "Hartmut Sonnenberg"
            dicToAddressOnBehalf.Add "address", m_UserEmailOnBehalf
        Next lng
    
        'in wessen Auftrag?
        Set colSenderAs = New Collection
        dicMessage.Add "from", colSenderAs
        For lng = 1 To m_SenderOnBehalf.Count
            Set dicSenderAs = New Dictionary
            colSenderAs.Add dicSenderAs
            Set dicToAddressSenderAs = New Dictionary
            dicSenderAs.Add "emailAddress", dicToAddressSenderAs
            dicToAddressSenderAs.Add "name", "Test Outlook Versand"
            dicToAddressSenderAs.Add "address", m_SenderOnBehalf.Item(lng)
        Next lng
    
    
    ElseIf m_SenderType = enSenderSendAs Then
        Set colSenderAs = New Collection
        dicMessage.Add "from", colSenderAs
        For lng = 1 To m_SenderAs.Count
            Set dicSenderAs = New Dictionary
            colSenderAs.Add dicSenderAs
            Set dicToAddressSenderAs = New Dictionary
            dicSenderAs.Add "emailAddress", dicToAddressSenderAs
            dicToAddressSenderAs.Add "address", m_SenderAs.Item(lng)
        Next lng
    End If
    
    'Empfänger
    dicMessage.Add "toRecipients", colToRecipients
    For lng = 1 To m_Recipients.Count
        Set dicToRecipient = New Dictionary
        colToRecipients.Add dicToRecipient
        Set dicToEMailAddress = New Dictionary
        dicToRecipient.Add "emailAddress", dicToEMailAddress
        dicToEMailAddress.Add "address", m_Recipients.Item(lng)
    Next lng
    
    'Cc-Empfänger
    If Not m_CcRecipients Is Nothing Then
        If Not m_CcRecipients.Count = 0 Then
            Set colToCCRecipients = New Collection
            dicMessage.Add "ccRecipients", colToCCRecipients
            
            For lng = 1 To m_CcRecipients.Count
                Set dicToCCRecipient = New Dictionary
                colToCCRecipients.Add dicToCCRecipient
                Set dicToCCEMailAddress = New Dictionary
                dicToCCRecipient.Add "emailAddress", dicToCCEMailAddress
                dicToCCEMailAddress.Add "address", m_CcRecipients.Item(lng)
            Next lng
        End If
    End If
    
    'Bcc-Empfänger
    If Not m_BccRecipients Is Nothing Then
        If Not m_BccRecipients.Count = 0 Then
            
            Set colToBCCRecipients = New Collection
            dicMessage.Add "bccRecipients", colToBCCRecipients
            
            For lng = 1 To m_BccRecipients.Count
                Set dicToBCCRecipient = New Dictionary
                colToBCCRecipients.Add dicToBCCRecipient
                Set dicToBCCEMailAddress = New Dictionary
                dicToBCCRecipient.Add "emailAddress", dicToBCCEMailAddress
                dicToBCCEMailAddress.Add "address", m_BccRecipients.Item(lng)
            Next lng
        End If
    End If
    
    If m_Priority = enLow Then
        dicMessage.Add "importance", "low"
    ElseIf m_Priority = enNormal Then
        dicMessage.Add "importance", "normal"
    Else
        dicMessage.Add "importance", "high"
    End If
    
    dicMail.Add "saveToSentItems", "true"
    
    If Not m_Attachments Is Nothing Then
        If Not m_Attachments.Count = 0 Then
            
            Set colAttachments = New Collection
            
            dicMessage.Add "attachments", colAttachments
            For lng = 1 To m_Attachments.Count
                Set dicAttachments = New Dictionary
                colAttachments.Add dicAttachments
                dicAttachments.Add "@odata.type", "#microsoft.graph.fileattachment"
                dicAttachments.Add "name", GetFilename(m_Attachments.Item(lng))
                strFileContent = EncodeFileToBase64(m_Attachments.Item(lng))
                dicAttachments.Add "contentBytes", strFileContent
            Next lng
        End If
    End If
    
    strJSON = ConvertToJson(dicMail)
    
    'kleine Korrekturen am JSON-String
    If m_SenderType = enSenderSendAs Then
        
        strJSON = Replace(strJSON, """from"":[{", """from"":{")
        strJSON = Replace(strJSON, "}}],""toRecipients""", "}},""toRecipients""")
    
    ElseIf m_SenderType = enSenderOnBehalf Then
        strJSON = Replace(strJSON, """sender"":[{", """sender"":{")
        strJSON = Replace(strJSON, "}}],""from""", "}},""from""")
        
        strJSON = Replace(strJSON, """from"":[{", """from"":{")
        strJSON = Replace(strJSON, "}}],""Recipients""", "}},""Recipients""")
    End If
    
     
    m_JSON_Email = strJSON
    
    Save_JsonString False, "EmailCreate", strJSON
    
    Email_Create = True

Ende:
    Set dicBody = Nothing
    Set dicMessage = Nothing
    Set dicMail = Nothing
    
    'Versender on behalf
    Set colOnBehalf = Nothing
    Set dicSenderOnBehalf = Nothing
    Set dicToAddressOnBehalf = Nothing
   
   'Empfänger
    Set colToRecipients = Nothing
    Set dicToRecipient = Nothing
    Set dicToEMailAddress = Nothing
    
    'Versender Send As
    Set colSenderAs = Nothing
    Set dicSenderAs = Nothing
    Set dicToAddressSenderAs = Nothing
    
    'Dateianhänge
    Set colAttachments = Nothing
    Set dicAttachments = Nothing
    
    Exit Function
     
ErrHandler:
     fehler = Err.Description
     Resume Ende
End Function


'===============================================================================
' Name: Function Email_Send
' Input:
'   ByRef fehler As String      - Verwendung by Ref!
' Output:
'   True/False
' Purpose:  ruft die Email (JSON-Format) ab
'           bestimmt den spezifischen Endpunkt sendMail und ruft ihn per POST auf
'           sendet JSON Mail an den Endpunkt
' Remarks: Author H. Sonnenberg Januar 2026
'===============================================================================
Public Function Email_Send(fehler As String) As Boolean
    Dim httpReq As MSXML2.XMLHTTP60
    Dim strURL As String
    Dim strToken As String
    Dim strbody As String
    Dim strResponse As String
    Dim success As Boolean
    
    On Error GoTo ErrHandler
    
    
   ' strURL = "https://graph.microsoft.com/v1.0/users/2bfff104-c4c9-4c31-92f3-512f3e18aa60/sendMail"
  '  strURL = "https://graph.microsoft.com/v1.0/users/" & m_UserEmailOnBehalf & "/sendMail"
    strURL = "https://graph.microsoft.com/v1.0/me/sendMail"

    'HTML Email erzeugen
    strbody = m_JSON_Email

    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    httpReq.Open "POST", strURL, False
    httpReq.setRequestHeader "Authorization", "Bearer " & m_Token
    httpReq.setRequestHeader "Content-Type", "application/json"
    httpReq.send strbody
    
    Select Case httpReq.Status
        Case 202
            Save_JsonString False, "Email_Send", httpReq.responseText
            Email_Send = True
        Case Else
            Save_JsonString True, "Email_Send", httpReq.responseText
            MsgBox "Fehler beim Senden der E-Mail: " & httpReq.Status & " - " & httpReq.responseText
            Debug.Print httpReq.responseText
    End Select
     
     
Ende:
    Set httpReq = Nothing
    Exit Function
     
ErrHandler:
    fehler = Err.Description
    Save_JsonString True, "Email_SendErrHandler", fehler
    Resume Ende

End Function


'===============================================================================
' Name: Function Request_UserInfos
' Input:
'   ByRef fehler As String      - Verwendung by Ref!
' Output:
'   True/False
' Purpose:  fragt über MSGraph alle für die Email-Signatur wichtigen Attribute
'           des User Ressourcentyps ab. Das sind displayName, department, companyname,
'           streetaddress, postalCode, city, businessphones, mail. Die Angaben werden
'           in den geleichnamigen Properties für spätere Verwendung festgehalten.
' Remarks: Author H. Sonnenberg Januar/Februar 2026
'===============================================================================
Public Function Request_UserInfos(fehler As String) As Boolean
    
    Dim httpReq As New MSXML2.XMLHTTP60
    Dim userInfoURL As String, strJSON As String
    Dim objJSON As Object
    
    On Error GoTo ErrHandler
 
    userInfoURL = "https://graph.microsoft.com/v1.0/me?$select=displayName, department , companyName, streetAddress, postalCode, city, businessPhones, mobilephone, mail"
    
    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    httpReq.Open "GET", userInfoURL, False
    httpReq.setRequestHeader "Authorization", "Bearer " & m_Token
    httpReq.setRequestHeader "Content-Type", "application/json"
    
    httpReq.send
    
    Select Case httpReq.Status
        Case 200
            strJSON = httpReq.responseText
            Save_JsonString False, "userInfosNo1", strJSON
            strJSON = Replace(strJSON, "[", vbNullString)
            strJSON = Replace(strJSON, "]", vbNullString)
            Save_JsonString False, "userInfosNo2", strJSON
            Set objJSON = ParseJson(strJSON)
            
            m_DisplayName = machNullString(objJSON("displayName"))
            m_Department = machNullString(objJSON("department"))
            m_Companyname = machNullString(objJSON("companyName"))
            m_Streetaddress = machNullString(objJSON("streetAddress"))
            m_PostalCode = machNullString(objJSON("postalCode"))
            m_City = machNullString(objJSON("city"))
            m_Businessphones = machNullString(objJSON("businessPhones"))
            m_Mobilephone = machNullString(objJSON("mobilePhone"))
            m_Mail = machNullString(objJSON("mail"))
            
        Case Else
            Save_JsonString True, "userInfoserror", httpReq.responseText
            fehler = "Fehler beim Senden der E-Mail: " & httpReq.Status & " - " & httpReq.responseText
            'Debug.Print httpReq.responseText
    End Select
     
    If fehler = vbNullString Then Request_UserInfos = True
    
    
Ende:
    Set httpReq = Nothing
    
    Exit Function
ErrHandler:
    fehler = Err.Description
    Save_JsonString True, "userInfosErrHandler", fehler
    Resume Ende

End Function

'===============================================================================
' Name: Function Request_Token
' Input:
'   ByRef fehler As String      - Verwendung by Ref!
' Output:
'   True/False
' Purpose:  fordert unter Verwendung der Client-ID der App, des Autorisierungscodes
'           der Berechtigungen ein Token an
'           das Token und auch das Refreshtoken werden in den klasseninternen
'           Variablen m_Token und m_RefreshToken gespeichert. Auch die Gültigkeitsdauer
'           wird in m_ExpiresInValue festgehalten
' Remarks: Author H. Sonnenberg Januar 2026
'===============================================================================
Public Function Request_Token(fehler As String) As Boolean
    Dim httpReq As MSXML2.XMLHTTP60
    Dim tokenURL As String, success As Boolean
    Dim objJSON As Object
    Dim strURL As String
    Dim strData As String
    Dim strJSON As String
    Dim strToken As String
    
    On Error GoTo ErrHandler
    
    strURL = "https://login.microsoftonline.com/" & cTenant_ID & "/oauth2/v2.0/token"
    strData = "grant_type=authorization_code" & _
           "&client_id=" & m_ClientID & _
           "&code=" & m_Authorization_Code & _
           "&scope=" & m_Scope & " " & _
           "&redirect_uri=" & cRedirect_URL

    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    httpReq.Open "POST", strURL, False
    httpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    httpReq.send strData

    Select Case httpReq.Status
        Case 200
            strJSON = httpReq.responseText
            Set objJSON = ParseJson(strJSON)
            Request_Token = True
            m_Token = objJSON("access_token")
            If m_Use_RefreshToken Then
                m_RefreshToken = objJSON("refresh_token")
            End If
            m_ExpiresInValue = objJSON("expires_in")
        Case Else
            strJSON = httpReq.responseText
            Set objJSON = ParseJson(strJSON)
            fehler = objJSON.Item("error_description")
    End Select
    
Ende:
    Set httpReq = Nothing
    Set objJSON = Nothing
    Exit Function
    
ErrHandler:
    fehler = Err.Description
    Resume Ende
 
End Function


'===============================================================================
' Name: Function Request_Authorization
' Input:
'   ByRef fehler As String      - Verwendung by Ref!
' Output:
'   True/False
' Purpose:  fordert unter Verwendung der Client-ID der App, des Response_Typs und
'           des Response_Mode den Autorisierungscode an
'           Dieser wird in der klasseninternen Variablen m_Authorization_Code gespeichert.
' Remarks: Author H. Sonnenberg Januar 2026
'===============================================================================
Public Function Request_Authorization(fehler As String) As Boolean
    'interessante Parameter des Requests
    'Prompt (Prompt=login oder prompt=none oder prompt=consent prompt=select_account)
        
    Dim httpReq As New MSXML2.XMLHTTP60
    Dim tokenURL As String, success As Boolean
    
    On Error GoTo ErrHandler
    
    'Authorization Code zurücksetzen
    m_Authorization_Code = vbNullString
    
    tokenURL = "https://login.microsoftonline.com/" & cTenant_ID & "/oauth2/v2.0/authorize? "
    tokenURL = tokenURL & "client_id=" & m_ClientID & " "
    tokenURL = tokenURL & "&response_type=" & m_Response_Type & ""      'siehe Class_Initialize
    tokenURL = tokenURL & "&redirect_uri=" & cRedirect_URL & " "
    tokenURL = tokenURL & "&response_mode=" & m_Response_Mode & " "     'siehe Class_Initialize

    tokenURL = tokenURL & "&scope=" & m_Scope & " "
    tokenURL = tokenURL & "&state=" & m_state                           'siehe Class_Initialize

    httpReq.Open "GET", tokenURL, False
    httpReq.setRequestHeader "Content-Type", "application/json"
    'httpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    httpReq.send
    If httpReq.Status = 200 Then
        If InStr(httpReq.responseText, "?code=") > 0 Then
            'genau dann haben wir den Authorisierungscode erhalten
            m_Authorization_Code = Get_AuthorizationCode(httpReq.responseText, success)
            Request_Authorization = True
        ElseIf InStr(httpReq.responseText, "Es wurde eine ung\u0026#252;ltige Anforderung empfangen") > 0 Then
            fehler = "Es wurde eine ungültige Anforderung empfangen" '&
        ElseIf InStr(httpReq.responseText, "Bei Ihrem Konto anmelden") > 0 Then
            fehler = "Bitte melden Sie sich bei Konto an" '"Bitte melden Sie sich bei Konto an (z.B. bei MS Teams)"
            
            On Error Resume Next
            tokenURL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?scope=service%3A%3Aaccount.microsoft.com%3A%3AMBI_SSL%20openid%20profile%20offline_access&response_"
            tokenURL = tokenURL & "type=code&client_id=81feaced-5ddd-41e7-8bef-3e20a2689bb7&redirect_uri=https%3A%2F%2Faccount.microsoft.com%2Fauth%2Fcomplete-signin-oauth"
            tokenURL = tokenURL & "&client-request-id=d11a5e89-de30-4f1c-8e14-ede2a1f12bfd&x-client-SKU=MSAL.Desktop&x-client-Ver=4.79.2.0&prompt=login&client_info=1"
            tokenURL = tokenURL & "&state=H4sIAAAAAAAEAAXBS4JDMAAA0Lt02wURrVgqaSkpIZjYKfWvT4cyTj_vHVxTx0eeztSzkLkXRYXDJH7c6U0dU7UWU-Dvo0cFcQfSMZ7XZn_Da6ussGVxaAl8g5harX7mptW3CGEvlTTdmvTk59KhZHEWCY48GOc4dv0CKntV2HdFRrLxlGzS-"
            tokenURL = tokenURL & "5 SVYJsilyVpx6YyKENRxA3kpbtUKL9g543O1reQ91eNsE29sKWL -n2OG1XFN4iDIBIu1kIyj5fB4yvkCqmNzNxuNOuMP3uow0c9sHhQtoX_Dk1O5FOVkchJ_PXj5sYHpS8YghOVep - RyLzh1XfWqhauK1KOWgt7FQLwq82NTDqxas_aE - k1UD0OQeLRIW"
            tokenURL = tokenURL & "KTfvgHMG-n7UIBAAA&msaoauth2=true&instance_aware=true&lc=1031"

            ActiveWorkbook.FollowHyperlink Address:=tokenURL       'Problem TestOutlook.de nicht erreichbar
            fehler = "Abbruch"
            
        ElseIf InStr(httpReq.responseText, "?code=") = 0 Then
            fehler = "der Authorisierungs-Code wurde nicht gesendet" '&
        End If
    End If


Ende:
    Set httpReq = Nothing
    
    Exit Function
ErrHandler:
    MsgBox Err.Description, ThisWorkbook.Name, Err.Description
    Resume Ende
End Function

'===============================================================================
' Name: Function LinkPermissions
' Input:
'   None
' Output:
'   True/False
' Purpose:  alle Berechtigungen aus der internen Collection m_Permissions werden
'           in die interne Variable m_scope überführt
' Remarks: Author H. Sonnenberg Januar 2026
'===============================================================================
Public Function LinkPermissions() As Boolean
    'alle Elemente der Collection m_Permissions zu einem String in Modulvariabler m_Scope verbinden!
    Dim ctr As Byte
    
    On Error GoTo ErrHandler
    
    If Not m_Permissions.Count = 0 Then
        If m_Use_RefreshToken Then
            'es kann ein abgelaufenes Token mit einem Aktualisierungstoken erneuert werden - siehe RefreshToken
            m_Scope = "offline_access"
        End If
        For ctr = 1 To m_Permissions.Count
            m_Scope = m_Scope & c_Blank & m_Permissions.Item(ctr)
        Next ctr
    End If
    
    LinkPermissions = True

    Exit Function
ErrHandler:
    MsgBox Err.Description, ThisWorkbook.Name, Err.Description
End Function

'===============================================================================
' Name: Function Get_AuthorizationCode
' Input:
'   ByRef httpRequest As String
'   ByRef success As Boolean
' Output:
'   String
' Purpose:  isoliert den Autorisierungscode aus dem httpReq.responseText
'           eird der internen Variablen m_Authorization_Code zugewiesen
' Remarks: Author H. Sonnenberg Januar 2026
'===============================================================================
Private Function Get_AuthorizationCode(httpRequest As String, success As Boolean) As String
    Dim posleft As Integer, posright As Integer
    
    On Error GoTo ErrHandler
    'nach dem Code folgt die Zeichenkette \u0026state=12345+HTTP%2f1.1 und das beginnt mit dem URL-Parameter Trennzeichen &, gefolgt von der Wiederholung des
    'in Request_Authorization() zufällig generierten eindeutiger Wert  "12345 HTTP/1.1"
    
    posleft = InStr(httpRequest, "?code=")
    posright = InStr(httpRequest, "\u0026state")
    
    Get_AuthorizationCode = Mid(httpRequest, posleft + Len("?code="), posright - (posleft + Len("?code=")))
    
    success = True
    
    Exit Function
ErrHandler:
    MsgBox Err.Description, ThisWorkbook.Name, Err.Description
End Function


'===============================================================================
' Name: Function GetFilename
' Input:
'   ByRef strPath As String
' Output:
'   String
' Purpose:  isoliert aus einer Datei- und Pfadangabe den Dateinamen
' Remarks: Author H. Sonnenberg Januar 2026
'===============================================================================
Private Function GetFilename(strPath As String) As String
    Dim strFilename As String
    strFilename = Mid(strPath, InStrRev(strPath, "\") + 1)
    GetFilename = strFilename
End Function

'===============================================================================
' Name: Function EncodeFileToBase64
' Input:
'   ByRef strPath As String
' Output:
'   String
' Purpose:  Binärdaten werden mit der Base64-Codierung in Text konvertiert
' Remarks: Author H. Sonnenberg Januar 2026
'===============================================================================
Function EncodeFileToBase64(strPath As String) As String
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As IXMLDOMElement
    Dim objInputStream As Object
    Dim bytData() As Byte
    Dim strBase64 As String

    Set objInputStream = CreateObject("ADODB.Stream")
    objInputStream.Type = 1 ' Binärdaten
    objInputStream.Open
    objInputStream.LoadFromFile strPath
    bytData = objInputStream.Read
    objInputStream.Close

    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("base64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = bytData
    strBase64 = Replace(objNode.Text, vbLf, "")
    
    EncodeFileToBase64 = strBase64
End Function


'******************************************************************************************************************************************
'************************************************           Events           **************************************************************
'******************************************************************************************************************************************

'===============================================================================
' Name: Sub Class_Initialize
' Input:
'   None
' Output:
'   None
' Purpose:  notwendige Collections werden initialisiert, interne Variablen mit
'           unveränderlichen Werten Versehen
'           wird automatisch mit der Instantiierung der Klasse aufgerufen
' Remarks: Author H. Sonnenberg Januar 2026
'===============================================================================
Private Sub Class_Initialize()
    Set m_Permissions = New Collection
    Set m_Recipients = New Collection
    Set m_SenderOnBehalf = New Collection
    Set m_SenderAs = New Collection
    
    m_Response_Type = "code"
    m_Response_Mode = "query"
    m_state = "12345 HTTP/1.1"
End Sub

'===============================================================================
' Name: Sub Class_Terminate
' Input:
'   None
' Output:
'   None
' Purpose:  allen initialisierten Collections wird Nothing zugewiesen
'           wird automatisch mit der Terminierung der Klasse aufgerufen
' Remarks: Author H. Sonnenberg Januar 2026
'===============================================================================
Private Sub Class_Terminate()
    Set m_Permissions = Nothing
    If Not m_SenderAs Is Nothing Then Set m_SenderAs = Nothing
    If Not m_SenderOnBehalf Is Nothing Then Set m_SenderOnBehalf = Nothing
    If Not m_Recipients Is Nothing Then Set m_Recipients = Nothing
    If Not m_CcRecipients Is Nothing Then Set m_CcRecipients = Nothing
    If Not m_BccRecipients Is Nothing Then Set m_BccRecipients = Nothing
    If Not m_Attachments Is Nothing Then Set m_Attachments = Nothing
End Sub




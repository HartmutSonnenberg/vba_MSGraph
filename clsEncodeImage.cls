VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEncodeImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name: clsEncodeImage
' Purpose:      Klasse beinhaltet Properties und Methoden, um Binärdateien
'               in das Base64 Format zu konvertieren
'               Besonderheit hier, die Klasse kann eine Bilddatei aus der Angabe
'               des Worksheet- und Chartnamen einer Exceldatei erzeugen
'               Bilddatei kann aber auch direkt übergeben werden
' Remarks:      Author H. Sonnenberg im Januar/Februar 2026
' Functions:    SaveChartToFile, FindWorkSheet, FindChart, EncodeFile
' Properties:   nameWorkSheet, DirectoryToSave, ChartName, Picture
' Methods:
' Started:      27.12.2025
' Modified:
'===============================================================================
Option Explicit

Private m_NameWorksheet As String                               'Worksheet mit Chartobjekt
Private m_DirectoryToSave As String                             'Zielverzeichnis zum Speichern des Images
Private m_ChartName As String                                   'Name des Chartobjektes im Worksheet
Private m_PictureFile As String                                 'wenn eine Bilddatei bereits vorhanden ist

'******************************************************************************************************************************************
'************************************************          Properties          ************************************************************
'******************************************************************************************************************************************

'write only
Public Property Let nameWorkSheet(strNameWorkSheet As String)
    m_NameWorksheet = strNameWorkSheet
End Property
'write only
Public Property Let DirectoryToSave(strDirToSave As String)
    m_DirectoryToSave = strDirToSave
End Property
'write only
Public Property Let chartName(strNameChart As String)
    m_ChartName = strNameChart
End Property
'wird nur verwendet, wenn eine Bilddatei lediglich angegeben werden muss
Public Property Let Picture(strPictureFile As String)
    m_PictureFile = strPictureFile
End Property



'******************************************************************************************************************************************
'************************************************          Methoden          **************************************************************
'******************************************************************************************************************************************
'===============================================================================
' Name: Sub SaveChartToFile
' Input:
'   ByRef savedFile As String   - Verwendung by Ref!
'   ByRef fehler As String      - Verwendung by Ref!
' Output:
'   True/False
' Purpose:  identifiziert Chart und exportiert es als png-Datei
'           liefert Vorarbeit zur Konvertierung in die Base64 Kodierung
' Remarks: Author H. Sonnenberg Januar 2026
'===============================================================================
Public Function SaveChartToFile(savedFile As String, fehler As String) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sh As Shape
    Dim cha As Chart
    Dim fileName As String
    
    Dim byteError As Byte
    
    On Error GoTo ErrHandler
    
    If FindWorkSheet(ActiveWorkbook, m_NameWorksheet) = False Then
        byteError = 1
        Err.Raise 901, "SaveChartToFile", "Worksheet """ & m_NameWorksheet & """ nicht vorhanden!"
    End If
    
    If FindChart(ActiveWorkbook, m_NameWorksheet, m_ChartName) = False Then
        byteError = 2
        Err.Raise 901, "SaveChartToFile", "Chart """ & m_ChartName & """ im Worksheet """ & m_NameWorksheet & """ nicht gefunden!"
    End If
    
    Set sh = ActiveWorkbook.Worksheets(m_NameWorksheet).Shapes(m_ChartName)
    
    fileName = m_DirectoryToSave & "\ActiveChart_C" & sh.TopLeftCell.Column & "_R" & sh.TopLeftCell.row & ".png"
    If Dir(fileName, vbNormal) <> "" Then Kill fileName
    Set cha = sh.Chart
    cha.Export fileName:=fileName, FilterName:="png"

    savedFile = fileName
    SaveChartToFile = True
    
    Exit Function
    
ErrHandler:
    fehler = Err.Description
End Function


'===============================================================================
' Name: function FindWorkSheet
' Input:
'   ByRef wb As Workbook
'   ByRef fehler As String      - Verwendung by Ref!
' Output:
'   True/False
' Purpose: prüft, ob angegebenes Worksheet vorhanden ist
' Remarks: Author H. Sonnenberg Januar 2026
'===============================================================================
Private Function FindWorkSheet(ByVal wb As Workbook, fehler As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    
    Set ws = wb.Worksheets(m_NameWorksheet)
    
    If Err.Number = 0 Then
        FindWorkSheet = True
    Else
        Err.Clear
    End If
    
    Set ws = Nothing
    
End Function

'===============================================================================
' Name: function FindChart
' Input:
'   ByRef wb As Workbook
'   ByRef nameWorkSheet As String
'   ByRef nameChart As String
' Output:
'   True/False
' Purpose: prüft, ob angegebenes Chart vorhanden ist
' Remarks: Author H. Sonnenberg Januar 2026
'===============================================================================
Private Function FindChart( _
            ByVal wb As Workbook, _
            ByVal nameWorkSheet As String, _
            ByVal nameChart As String) As Boolean
            
    Dim sh As Shape
    Dim cha As Chart
    
    On Error Resume Next
    
    Set sh = wb.Worksheets(nameWorkSheet).Shapes(nameChart)
    
    If Err.Number = 0 Then
        FindChart = True
    Else
        Err.Clear
    End If
    
    
    Set sh = Nothing
    
End Function

'===============================================================================
' Name: function EncodeFile
' Input:
'   ByRef imagePath As String
'   ByRef imagePath As String
'   ByRef fehler As String
' Output:
'   True/False
' Purpose:  eine Binärdatei (hier Image Argument imagepath) wird Base64 kodiert
'           und per Referenz "zurückgegeben"
' Remarks: Author H. Sonnenberg Januar 2026
'===============================================================================
Public Function EncodeFile(imagePath As String, encodedFile As String, fehler As String) As Boolean
    Dim objXML As Object
    Dim objDocElem As Object
    Dim objStream As Object
    
    Const adTypeBinary = 1          ' Binary file is encoded
    
    On Error GoTo ErrHandler
 
    'Stream Objekt für das Image
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = adTypeBinary
    objStream.Open
    objStream.LoadFromFile (imagePath)
    
    'XMLDocument und Wurzelknoten mit den Base64 Daten
    Set objXML = CreateObject("MSXml2.DOMDocument")
    Set objDocElem = objXML.createElement("Base64Data")
    objDocElem.DataType = "bin.base64"
    
    'Base64 Daten einlesen und zuweisen
    objDocElem.nodeTypedValue = objStream.Read()
        
    'Base64 Daten zuweisen
    encodedFile = objDocElem.Text
    EncodeFile = True
 
Ende:
    
    Set objXML = Nothing
    Set objDocElem = Nothing
    Set objStream = Nothing
    
    Exit Function
    
ErrHandler:
    fehler = Err.Description
    Resume Ende
End Function



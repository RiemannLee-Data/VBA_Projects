VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "klsDerivat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Derivat Klasse Modul

'Mitglied Variablen
Private s_referenz As String
Private s_spalte As Long
Private s_typschluessel As String
Private s_name As String
Private s_gueltigkeitsdatum As String


'Eigenschaften
Public Property Get referenz() As String
    referenz = s_referenz
End Property



Public Property Get Spalte() As Long
    Spalte = s_spalte
End Property



Public Property Get Typschluessel() As String
    Typschluessel = s_typschluessel
End Property



Public Property Get Name() As String
    Name = s_name
End Property



Public Property Get Gueltigkeitsdatum() As String
    Gueltigkeitsdatum = s_gueltigkeitsdatum
End Property



Public Property Let referenz(ByVal s_referenz_new As String)
    s_referenz = s_referenz_new
End Property



Public Property Let Spalte(ByVal s_spalte_new As Long)
    s_spalte = s_spalte_new
End Property



Public Property Let Typschluessel(ByVal s_typschluessel_new As String)
    s_typschluessel = s_typschluessel_new
End Property



Public Property Let Name(ByVal s_name_new As String)
    s_name = s_name_new
End Property



Public Property Let Gueltigkeitsdatum(ByVal s_gueltigkeitsdatum_new As String)
    s_gueltigkeitsdatum = s_gueltigkeitsdatum_new
End Property



'Methoden
Public Sub Init(ws1 As Worksheet, ws2 As Worksheet, ref As String)
'
'Beschreibung: Initialisiert Derivate von LU Strukturreport
'
    s_referenz = ref
    s_spalte = ws2.Cells.Find(What:=ref, LookAt:=xlWhole).Column
    s_typschluessel = ws2.Cells.Find(What:=ref, LookAt:=xlWhole).Offset(rowOffset:=1, columnOffset:=0).Value
    s_name = Split(ref, " ")(0)
    s_gueltigkeitsdatum = CStr(Replace(ws1.Cells.Find(What:="Datum", LookAt:=xlWhole).Offset(rowOffset:=0, columnOffset:=1), ".", ""))
End Sub



Public Sub Speichern(ws As Worksheet)
'
'Beschreibung: Schreibt gewählte Derivat Daten in LOG Tabelle
'
    ws.Cells(1, 2).Value = s_referenz
    ws.Cells(2, 2).Value = s_spalte
    ws.Cells(3, 2).Value = s_typschluessel
    ws.Cells(4, 2).Value = s_name
    ws.Cells(5, 2).Value = s_gueltigkeitsdatum
    
End Sub



Public Sub Lesen(ws As Worksheet)
'
'Beschreibung: Liest gespeicherte Daten von gewählte Derivat in LOG Tabelle
'
    s_referenz = ws.Cells(1, 2).Value
    s_spalte = ws.Cells(2, 2).Value
    s_typschluessel = ws.Cells(3, 2).Value
    s_name = ws.Cells(4, 2).Value
    s_gueltigkeitsdatum = CStr(ws.Cells(5, 2).Value)
End Sub



Public Sub InitStandardExport(ws1 As Worksheet, ws2 As Worksheet, ref As String)
'
'Beschreibung: Initialisiert Derivate von Standard Strukturbericht Export aus PDM (für Benutzung in LK/LG/LI PL)
'
    s_referenz = ref
    s_spalte = ws2.Cells.Find(What:=ref, LookAt:=xlWhole).Column
    s_typschluessel = Split(ref, " ")(0)
    s_name = Split(ref, " ")(1)
    s_gueltigkeitsdatum = CStr(Replace(ws1.Cells.Find(What:="Datum", LookAt:=xlWhole).Offset(rowOffset:=0, columnOffset:=1), ".", ""))

End Sub

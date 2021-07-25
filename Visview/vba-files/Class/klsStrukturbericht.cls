VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "klsStrukturbericht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Strukturbericht Klasse Modul

Private s_kopfzeile As Long

Private s_fbspalte As Long
Private s_modulorgspalte As Long
Private s_kogrspalte As Long
Private s_komponentespalte As Long
Private s_guidspalte As Long

Private s_zeilenanzahl As Long
Private s_spaltenanzahl As Long
Private s_name As String



Public Property Get Kopfzeile() As Long
    Kopfzeile = s_kopfzeile
End Property



Public Property Let Kopfzeile(ByVal s_kopfzeile_new As Long)
    s_kopfzeile = s_kopfzeile_new
End Property



Public Property Get FBSpalte() As Long
    FBSpalte = s_fbspalte
End Property



Public Property Let FBSpalte(ByVal s_fbspalte_new As Long)
    s_fbspalte = s_fbspalte_new
End Property



Public Property Get ModulOrgSpalte() As Long
    ModulOrgSpalte = s_modulorgspalte
End Property



Public Property Let ModulOrgSpalte(ByVal s_modulorgspalte_new As Long)
    s_modulorgspalte = s_modulorgspalte_new
End Property



Public Property Get KoGrSpalte() As Long
    KoGrSpalte = s_kogrspalte
End Property



Public Property Let KoGrSpalte(ByVal s_kogrspalte_new As Long)
    s_kogrspalte = s_kogrspalte_new
End Property



Public Property Get KomponenteSpalte() As Long
    KomponenteSpalte = s_komponentespalte
End Property



Public Property Let KomponenteSpalte(ByVal s_komponentespalte_new As Long)
    s_komponentespalte = s_komponentespalte_new
End Property



Public Property Get GUIDSpalte() As Long
    GUIDSpalte = s_guidspalte
End Property



Public Property Let GUIDSpalte(ByVal s_guidspalte_new As Long)
    s_guidspalte = s_guidspalte_new
End Property



Public Property Get Zeilenanzahl() As Long
    Zeilenanzahl = s_zeilenanzahl
End Property



Public Property Let Zeilenanzahl(ByVal s_zeilenanzahl_new As Long)
    s_zeilenanzahl = s_zeilenanzahl_new
End Property



Public Property Get Spaltenanzahl() As Long
    Spaltenanzahl = s_spaltenanzahl
End Property



Public Property Let Spaltenanzahl(ByVal s_spaltenanzahl_new As Long)
    s_spaltenanzahl = s_spaltenanzahl_new
End Property



Public Property Get Name() As String
    Name = s_name
End Property



Public Property Let Name(ByVal s_name_new As String)
    s_name = s_name_new
End Property




Public Sub Init(ws As Worksheet)
    s_kopfzeile = ws.Cells.Find(What:="#", LookAt:=xlWhole).Row
    s_fbspalte = ws.Cells.Find(What:="FB", LookAt:=xlWhole).Column
    s_modulorgspalte = ws.Cells.Find(What:="Modulorg.", LookAt:=xlWhole).Column
    s_kogrspalte = ws.Cells.Find(What:="PPG", LookAt:=xlWhole).Column
    s_komponentespalte = ws.Cells.Find(What:="Komponente", LookAt:=xlWhole).Column
    s_guidspalte = ws.Cells.Find(What:="Knoten-GUID", LookAt:=xlWhole).Column
    
    s_zeilenanzahl = ws.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    s_spaltenanzahl = ws.Cells.Find("*", searchorder:=xlByColumns, searchdirection:=xlPrevious).Column
    s_name = ws.Name
End Sub
    
   
  

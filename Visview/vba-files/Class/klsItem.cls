VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "klsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Item Klassemodule

Private s_fachbereich As String
Private s_modulorg As String
Private s_kogr As String
Private s_komponente As String
Private s_kommunalitaet As String
Private s_guid As String



Public Property Get Fachbereich() As String
    Fachbereich = s_fachbereich
End Property



Public Property Let Fachbereich(ByVal s_fachbereich_new As String)
    s_fachbereich = s_fachbereich_new
End Property



Public Property Get ModulOrg() As String
    ModulOrg = s_modulorg
End Property



Public Property Let ModulOrg(ByVal s_modulorg_new As String)
    s_modulorg = s_modulorg_new
End Property



Public Property Get kogr() As String
    kogr = s_kogr
End Property



Public Property Let kogr(ByVal s_kogr_new As String)
    s_kogr = s_kogr_new
End Property



Public Property Get Komponente() As String
    Komponente = s_komponente
End Property



Public Property Let Komponente(ByVal s_komponente_new As String)
    s_komponente = s_komponente_new
End Property



Public Property Get Kommunalitaet() As String
    Kommunalitaet = s_kommunalitaet
End Property



Public Property Let Kommunalitaet(ByVal s_kommunalitaet_new As String)
    s_kommunalitaet = s_kommunalitaet_new
End Property



Public Property Get GUID() As String
    GUID = s_guid
End Property


Public Property Let GUID(ByVal s_guid_new As String)
    s_guid = s_guid_new
End Property



Public Sub Init(ws As Worksheet, oDerivat As klsDerivat, oStrukturbericht As klsStrukturbericht, index As Long)
    s_fachbereich = ws.Cells(index, oStrukturbericht.FBSpalte)
    s_modulorg = ws.Cells(index, oStrukturbericht.ModulOrgSpalte)
    s_kogr = ws.Cells(index, oStrukturbericht.KoGrSpalte)
    s_komponente = ws.Cells(index, oStrukturbericht.KomponenteSpalte)
    s_guid = ws.Cells(index, oStrukturbericht.GUIDSpalte)
    
    s_kommunalitaet = ws.Cells(index, oDerivat.Spalte)
    
End Sub


Function Valid() As Boolean
    Let Valid = (s_kommunalitaet = "g") Or (s_kommunalitaet = "n") Or (s_kommunalitaet = "s")
End Function

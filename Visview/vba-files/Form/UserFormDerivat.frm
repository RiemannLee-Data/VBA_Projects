VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormDerivat 
   Caption         =   "Wählen Sie ein Derivat"
   ClientHeight    =   3220
   ClientLeft      =   -10
   ClientTop       =   220
   ClientWidth     =   5640
   OleObjectBlob   =   "UserFormDerivat.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserFormDerivat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'UserFormDerivat Modul

Private Sub UserForm_Initialize()

    Dim der As Variant
    Dim j As Integer
    der = Split(derivat_liste, ", ")
    For j = LBound(der) To UBound(der) - 1
        ComboBoxDerivat.AddItem Split(derivat_liste, ", ")(j)
    Next j
    
    FillTypCombo ComboBoxDerivat.Value
    
End Sub

Private Sub ComboBoxDerivat_Change()

    FillTypCombo ComboBoxDerivat.Value
    
End Sub

Sub FillTypCombo(ByVal val As String)

    Dim k As Integer
    Dim der As klsDerivat
    ComboBoxTyp.Clear

    For Each der In derivat_sammlung
        If der.Name = val Then
            ComboBoxTyp.AddItem der.referenz
        End If
    Next der
    
End Sub

Private Sub WaehlenSchaltflaeche1_Click()
    Dim der As New klsDerivat
    
    For Each der In derivat_sammlung
        If der.referenz = ComboBoxTyp.Value Then
            gewaehlt = der.Spalte
        End If
    Next der

    Hide
    
End Sub

'Private Sub WaehlenSchaltflaeche1_Enter()
'    Dim der As New klsDerivat
    
'    For Each der In derivat_sammlung
 '       If der.Referenz = ComboBoxTyp.Value Then
'            gewaehlt = der.Spalte
'        End If
'    Next der

 '   Hide
'End Sub

Private Sub BeendenSchaltflaeche1_Click()

    Unload Me
    
End Sub

'Private Sub BeendenSchaltflaeche1_Enter()

'    Unload Me
    
'End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormErzeugnis 
   Caption         =   "Wählen Sie ein Erzeugnis Format"
   ClientHeight    =   3200
   ClientLeft      =   0
   ClientTop       =   220
   ClientWidth     =   3160
   OleObjectBlob   =   "UserFormErzeugnis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormErzeugnis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'UserFormErzeugnis Modul

Private Sub AbbrechenSchaltflaeche_Click()
    End
End Sub

'Private Sub AbbrechenSchaltflaeche_Enter()
'    End
'End Sub

Private Sub EinzigeDateiOption_Click()
    erzeugnis_format = 1
End Sub

Private Sub FachbereichDateiOption_Click()
    erzeugnis_format = 2
End Sub

Private Sub OKSchaltflaeche_Click()
    Hide
End Sub

'Private Sub OKSchaltflaeche_Enter()
'    Hide
'End Sub
Private Sub UserForm_Click()

End Sub

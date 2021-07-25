Attribute VB_Name = "Visualisierung_Steuerelement"
Sub btn_weiter_Klicken()
    Call Dateierstellung2(der, Gültigkeitsdatum)
    Call Löschen
    For Each Sheet In ActiveWorkbook.Worksheets
        If Sheet.Name <> "Home" Then
        Sheet.Visible = False
        End If
    Next Sheet
    Call Shell("Explorer /e, " & strFile1, vbNormalFocus)
'    MsgBox ("XML-Daten wurden erstellt und im Ordner \Visualisierung_" & Derivat & " gespeichert.")

End Sub



Sub btn_abbrechen_Klicken()
    Call Löschen
        For Each Sheet In ActiveWorkbook.Worksheets
            If Sheet.Name <> "Home" Then
            Sheet.Visible = False
            End If
        Next Sheet
End Sub

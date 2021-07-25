Attribute VB_Name = "Alle_Daten_Loeschen"
'Alle_Daten_Loeschen Modul

Sub InhaltLoeschen()
'
'Beschreibung: Löscht Inhalt in PIVOT oder MAKRO Tabelle
'
    Dim Zeilenanzahl As Integer
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    
    Set ws1 = ThisWorkbook.Worksheets(1)
    Set ws2 = ThisWorkbook.Worksheets(2)
    ws1.Activate
    
    'Löscht MAKRO Tabelle Inhalt
    Zeilenanzahl = ws1.Cells(Rows.Count, 1).End(xlUp).Row
    ws1.Activate
    ws1.Range(Cells(2, 1), Cells(Zeilenanzahl + 100, 6)).Select
    Selection.Delete Shift:=xlUp
    
    'Löscht PIVOT Tabelle Inhalt
    ws2.Activate
    Zeilenanzahl = ws2.Cells(Rows.Count, 1).End(xlUp).Row
'    ws2.Range(Cells(2, 24), Cells(Zeilenanzahl + 100, 24)).Select
'    Selection.Delete Shift:=xlUp
    ws2.Range(Cells(2, 1), Cells(Zeilenanzahl + 100, 10)).Select
    Selection.Delete Shift:=xlUp
'    ws2.Range(Cells(2, 23), Cells(Zeilenanzahl + 100, 23)).Select
'    Selection.Delete Shift:=xlUp
'    ws2.Range(Cells(2, 7), Cells(Zeilenanzahl + 100, 7)).Select
'    Selection.Delete Shift:=xlUp
    
    'Formatieren
    ws2.Columns("A:Z").Select
    'Selection.Columns.AutoFit
    
    Selection.ColumnWidth = 10.71
    ws2.Cells(2, 1).Select
    
    Call RahmenZellenLoeschen(ThisWorkbook.Worksheets("LOG"), 1, 2, 8, 2)
    
    ws2.Activate
    
End Sub


Sub RahmenZellenLoeschen(ws As Worksheet, anfang_zeile As Long, anfang_spalte As Long, end_zeile As Long, end_spalte As Long)
    ws.Activate
    ws.Range(Cells(anfang_zeile, anfang_spalte), Cells(end_zeile, end_spalte)).Select
    Selection.Delete Shift:=xlUp
    'Selection.Columns.AutoFit
    
End Sub

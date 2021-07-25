Attribute VB_Name = "Daten_Analysieren"
'Analyse Modul

Public derivat_sammlung As Collection
Public derivat_liste As String
Public produkt_linie As Integer
Public gewaehlt As Integer
'
' Beschreibung:
'   here we make use of all the sub-functions to complete the final analyse
Sub DatenAnalysieren()

    Dim wk1 As Workbook
    Dim sh1 As Worksheet
    Dim sh2 As Worksheet
    Dim pivot As Worksheet
    Dim makro As Worksheet
    Dim log As Worksheet
    
    Dim StrukturberichtMappe As New klsStrukturbericht
    
    Dim itemColl As Collection
    
    Dim ufd As UserFormDerivat
    Dim ufpl As UserFormProduktLinie
    
    Dim derivatGewaehlt As New klsDerivat
    
    Dim Datei As Variant
    
    Dim cl_nummer As Integer
    
    'Applikation Eigenschaften
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AddCustomList Array("EA", "EE", "EF", "EP", "EV")
        
    'Arbeitsmappe Tabelle identifizieren
    Set wk1 = ThisWorkbook
    Set makro = wk1.Worksheets("MAKRO")
    Set pivot = wk1.Worksheets("PIVOT")
    Set log = wk1.Worksheets("LOG")
    
    'Produktlinie wählen
    Set ufpl = New UserFormProduktLinie
    ufpl.Show
    
    'Ausleitung Daten
    Let Datei = AnleitungWaehlen()
    Set wk2 = Workbooks.Open(Datei)
    Set sh1 = wk2.Worksheets("Kopf mit Parameter")
    Set sh2 = wk2.Worksheets("Strukturbericht")
    If sh1.FilterMode Then sh1.ShowAllData
    If sh2.FilterMode Then sh2.ShowAllData
        
    'Strukturbericht Tabelle Daten
    StrukturberichtMappe.Init sh2
    Call FormatTabelle(StrukturberichtMappe, sh2)
    
    'Wählt welche Derivat zu darstellen
    'Für LU Produktlinie
    If produkt_linie = 1 Then
        derivat_liste = ""
        Set derivat_sammlung = DerivatSammeln(sh1, sh2, StrukturberichtMappe)
    ElseIf produkt_linie = 2 Then
        derivat_liste = ""
        Set derivat_sammlung = DerivatSammelnStandardExport(sh1, sh2, StrukturberichtMappe)
    End If
    
    'Zeigt Wahlformular für den Derivaten
    Set ufd = New UserFormDerivat
    ufd.Show
    
    'Speichert die Daten den gewählte Derivat in LOG Tabelle
    If produkt_linie = 1 Then
        derivatGewaehlt.Init sh1, sh2, sh2.Cells(StrukturberichtMappe.Kopfzeile, gewaehlt)
    ElseIf produkt_linie = 2 Then
        derivatGewaehlt.InitStandardExport sh1, sh2, sh2.Cells(StrukturberichtMappe.Kopfzeile, gewaehlt)
    End If
    derivatGewaehlt.Speichern log
    
    'Markiert relevanten Daten in Ausleitung und speichert sie in eine Sammlung
    Set itemColl = KMGItemsSammeln(sh2, derivatGewaehlt, StrukturberichtMappe)
    
    'Schließt KMG Arbeitsmappe ohne Änderungen
    wk2.Close savechanges:=False
        
    'Schreibt die Sammlung von KMG Daten in Makro Tabelle
    Call KMGItemsSchreiben(makro, itemColl)
    
    'Vorverarbeitet die Daten
    Call Vorverarbeiten(makro)
    
    'Stellt die Pivot Tabelle dar in PIVOT Tabelle
    Call PivotErzeugen(makro, pivot)
    
    'Erzeugt die Regel für jede KoGr
    Call RegelErzeugen(pivot)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    cl_nummer = Application.GetCustomListNum(Array("EA", "EE", "EF", "EP", "EV"))
    Application.DeleteCustomList cl_nummer
    
End Sub

Function AnleitungWaehlen() As Variant

' Beschreibung:
'   Wahlformular der KMG Anleitung

    Dim Name As Variant
    
    '1. Öffnet Fenster wo der Benutzer wählt das Dokument
    '2. Speichert Pfad wenn der Benutzer auf "OK" Schaltfläche klickt
    Name = Application.GetOpenFilename(Title:="Wählen Sie bitte eine Ausleitung aus")
    
    'Wenn der Benutzer auf "Abbrechen" Schaltfläche klickt...
    If Name = False Then
        '... beendet das Makro
        End
    End If
    
    Let AnleitungWaehlen = Name
    
End Function



Function DerivatSammeln(ws1 As Worksheet, ws2 As Worksheet, oStrukturbericht As klsStrukturbericht) As Collection

' Beschreibung:
' Sammelt alle Derivaten die sind in KMG Strukturbericht gelistet
' Die Zellen sind erkennbar durch AWD, FWD, LL, RL Stichwörter.

    Dim zeile As Long
    Dim spalten As Long
    Dim anzahl As Integer
    Dim i As Integer
    Dim der As klsDerivat
    
    Let zeile = oStrukturbericht.Kopfzeile
    Let spalten = oStrukturbericht.Spaltenanzahl
    Set DerivatSammeln = New Collection
    
    For i = 1 To spalten
        anzahl = 0
        If InStr(1, ws2.Cells(zeile, i), "AWD") Then
            anzahl = anzahl + 1
        ElseIf InStr(1, ws2.Cells(zeile, i), "FWD") Then
            anzahl = anzahl + 1
        End If
        If InStr(1, ws2.Cells(zeile, i), "LL") Then
            anzahl = anzahl + 1
        ElseIf InStr(1, ws2.Cells(zeile, i), "RL") Then
            anzahl = anzahl + 1
        End If
        
        If anzahl = 2 Then
            If InStr(1, derivat_liste, Split(ws2.Cells(zeile, i), " ")(0)) < 1 Then
                If Split(ws2.Cells(zeile, i), " ")(0) = "John" Then
                    derivat_liste = derivat_liste & "John Cooper" & ", "
                Else
                    derivat_liste = derivat_liste & Split(ws2.Cells(zeile, i), " ")(0) & ", "
                End If
            End If
            Set der = New klsDerivat
            der.Init ws1, ws2, ws2.Cells(zeile, i)
            DerivatSammeln.Add der
        End If
    Next i
    
End Function



Function DerivatSammelnStandardExport(ws1 As Worksheet, ws2 As Worksheet, oStrukturbericht As klsStrukturbericht) As Collection

' Beschreibung:
'   Sammelt alle Derivaten, die in KMG Strukturbericht Export gelistet sind
'   Die Zellen sind erkennbar durch "KOMM_" in erste Zeile von Export (nur für LK;LG;LI)

    Dim zeile As Long
    Dim spalten As Long
    Dim anzahl As Integer
    Dim i As Integer
    Dim der As klsDerivat
    Dim anfang_range As Range
    Dim referenz As String
    
    Let zeile = oStrukturbericht.Kopfzeile
    Let spalten = oStrukturbericht.Spaltenanzahl
    Set DerivatSammelnStandardExport = New Collection
    
    Set anfang_range = ws2.Cells.Find(What:="KOMM_", LookAt:=xlPart)
    For i = anfang_range.Column To spalten
        If InStr(1, ws2.Cells(anfang_range.Row, i), "KOMM_") > 0 Then
            'Derivat listen
            If InStr(1, derivat_liste, Split(ws2.Cells(zeile, i), " ")(1)) < 1 Then
                derivat_liste = derivat_liste & Split(ws2.Cells(zeile, i), " ")(1) & ", "
            End If
            
            'Typen in eine Sammlung speichern
            Set der = New klsDerivat
            der.InitStandardExport ws1, ws2, ws2.Cells(zeile, i)
            DerivatSammelnStandardExport.Add der
        End If
    Next i
    
End Function



Function KMGItemsSammeln(ws As Worksheet, oDerivat As klsDerivat, oStrukturbericht As klsStrukturbericht) As Collection

' Beschreibung:
' Speichert relevanten Zeilen auf KMG Dokument in eine Kollektion

    Dim oItem As klsItem
    Dim zeile As Long

    Set KMGItemsSammeln = New Collection
    
    For zeile = 6 To oStrukturbericht.Zeilenanzahl
        Set oItem = New klsItem
        oItem.Init ws, oDerivat, oStrukturbericht, zeile
        If oItem.Valid Then
            KMGItemsSammeln.Add oItem
        End If
    Next zeile
    
End Function

Sub KMGItemsSchreiben(ws As Worksheet, oColl As Collection)

' Beschreibung:
' Schreibt die Item Kollektion in MAKRO Tabelle

    Dim i As Long
    Dim oItem As klsItem

    For i = 1 To oColl.Count
        Set oItem = oColl(i)
        ws.Cells(i + 1, "A") = oItem.Fachbereich
        ws.Cells(i + 1, "B") = oItem.ModulOrg
        ws.Cells(i + 1, "C") = oItem.kogr
        ws.Cells(i + 1, "D") = oItem.Kommunalitaet
        ws.Cells(i + 1, "E") = oItem.GUID
        ws.Cells(i + 1, "F") = oItem.Komponente
    Next i
    
End Sub



Sub Vorverarbeiten(ws As Worksheet)
''For debugging
'Sub Vorverarbeiten()
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Worksheets("MAKRO")

' Beschreibung:
' Vorverarbeitet Daten in MAKRO Tabelle
'   1. Kürzt PPG auf KoGr
'   2. Ändert Referenzen für einige Komponenten

    Dim i As Long
    Dim Zeilenanzahl As Long
    
    'MAKRO Tabelle Zeilenanzahl
    Let Zeilenanzahl = ws.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    
    'Löscht alle Zeilen mit "Kleinteile..." am Anfang der Komponente Zell
    'Löscht alle Zeilen mit "Formstücke..." am Anfang der Komponente Zell
    'Löscht alle Bordpapiere und Betriebsanleitung Komponenten
    For i = 2 To Zeilenanzahl
        Do While InStr(ws.Cells(i, 6), "Kleinteile") Or InStr(ws.Cells(i, 6), "Formstücke") Or ws.Cells(i, 2) = "KK01"
            ws.Rows(i).Delete
            Zeilenanzahl = Zeilenanzahl - 1
        Loop
    Next i
    
    For i = 2 To Zeilenanzahl
        'PPG-String kürzen auf KoGr
        ws.Cells(i, 3) = CStr(Format(Mid(ws.Cells(i, 3), 2, 4), "0000"))

        If IsEmpty(ws.Cells(i, 1)) And Not IsEmpty(ws.Cells(i, 6)) Then
            ws.Cells(i, 1) = "Leer <> EA"
            If IsEmpty(ws.Cells(i, 2)) And ws.Cells(i, 3) = "" Then
                ws.Cells(i, 3) = "Aus dem Motor"
            End If

            '"Generator..." in EA/MC07
            If InStr(1, ws.Cells(i, 6), "Generator") Then
                ws.Cells(i, 1) = "EA"
                ws.Cells(i, 2) = "MC07"
                ws.Cells(i, 3) = ""
            End If
            
            '"Riemen..." in EA/MD01
            If InStr(1, ws.Cells(i, 6), "Riemen mit") Then
                ws.Cells(i, 1) = "EA"
                ws.Cells(i, 2) = "MD01"
                ws.Cells(i, 3) = ""
            End If
            
            '"Zusatzwasserpumpe..." in EA/MD02
            If InStr(1, ws.Cells(i, 6), "Elektrische Zusatzwasserpumpe") Then
                ws.Cells(i, 1) = "EA"
                ws.Cells(i, 2) = "MD02"
                ws.Cells(i, 3) = ""
            End If
            
            '"Motorabdeckung..." in EA/MD06/1114
            If InStr(1, ws.Cells(i, 6), "Motorabdeckung") Then
                ws.Cells(i, 1) = "EA"
                ws.Cells(i, 2) = "MD06"
                ws.Cells(i, 3) = "1114"
            End If
            
        End If
    Next i
      
End Sub


' Beschreibung:
' Löscht Zeilenumbruch im Kopfzeile von Strukturbericht Mappe (nutzlich für Wahlform den Derivaten)

Sub FormatTabelle(oStrukturbericht As klsStrukturbericht, ws As Worksheet)

    For i = 1 To oStrukturbericht.Spaltenanzahl
        ws.Cells(oStrukturbericht.Kopfzeile, i) = Replace(ws.Cells(oStrukturbericht.Kopfzeile, i), Chr(10), " ")
    Next i
    
End Sub


' Beschreibung:
' wandelt die Daten aus der MAKRO Tabelle in ein PivotTable Objekt um und zeigt sie in der PIVOT Tabelle an

Sub PivotErzeugen(ws1 As Worksheet, ws2 As Worksheet)

    Dim Daten As String
    Dim Zeilenanzahl As Long
    Dim zeile As Range
    
    'MAKRO Tabelle Zeilenanzahl
    Let Zeilenanzahl = ws1.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    ws1.Activate

    ws1.Range("A2:F" & Zeilenanzahl).Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlNo
    
    ws2.Sort.SortFields.Clear
    
    'Markiert alle Daten von MAKRO Tabelle
    Let Daten = "MAKRO!R1C1:R" & Zeilenanzahl & "C6"
    
    'Schafft die Pivot Tabelle
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Daten, Version:=5).CreatePivotTable _
        TableDestination:="PIVOT!R2C1", TableName:="PivotTable", _
        DefaultVersion:=5
    
    ws2.Activate
    
    'Fachbereich Bereich als erste Ebene
    ws2.PivotTables("PivotTable").PivotFields("FB").Orientation = xlRowField
    ws2.PivotTables("PivotTable").PivotFields("FB").Position = 1
    
    'Modulorg Bereich als zweite Ebene
    ws2.PivotTables("PivotTable").PivotFields("ModulOrg").Orientation = xlRowField
    ws2.PivotTables("PivotTable").PivotFields("ModulOrg").Position = 2
        
    'KoGr Bereich als dritte Ebene
    ws2.PivotTables("PivotTable").PivotFields("KoGr").Orientation = xlRowField
    ws2.PivotTables("PivotTable").PivotFields("KoGr").Position = 3
    
    'Komponente GUID als vierte Ebene
    ws2.PivotTables("PivotTable").PivotFields("GUID").Orientation = xlRowField
    ws2.PivotTables("PivotTable").PivotFields("GUID").Position = 4
    
    'Komponente Bereich als fünfte Ebene
    ws2.PivotTables("PivotTable").PivotFields("Komponente").Orientation = xlRowField
    ws2.PivotTables("PivotTable").PivotFields("Komponente").Position = 5

    'Kommunalität Anzahl für jede Ebene rechnen
    ws2.PivotTables("PivotTable").AddDataField ws2.PivotTables("PivotTable").PivotFields("Treffer"), "Anzahl von Treffer", xlCount
    
    ws2.PivotTables("PivotTable").PivotFields("Treffer").Orientation = xlColumnField
    ws2.PivotTables("PivotTable").PivotFields("Treffer").Position = 1
    
    'Pivot-Table Style
    ws2.PivotTables("PivotTable").TableStyle2 = "PivotStyleMedium2"
    
    'Passt die Breite den Kommunalität Spalten
    ws2.Columns("B:D").ColumnWidth = 6
    
    ws2.PivotTables("PivotTable").PivotFields("FB").AutoSort xlManual, "FB"
    
    'PIVOT Tabelle Zeilenanzahl
    Let Zeilenanzahl = ws2.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    
    'Färbt alle Kommunalität Zellen mit der entsprechenden Farbe, die nicht leer sind
    For i = 4 To Zeilenanzahl
        If ws2.Cells(i, 2) <> "" Then
            ws2.Cells(i, 2).Interior.ColorIndex = 4 'grün
        End If
        If ws2.Cells(i, 3) <> "" Then
            ws2.Cells(i, 3).Interior.ColorIndex = 3 'rot
        End If
        If ws2.Cells(i, 4) <> "" Then
            ws2.Cells(i, 4).Interior.ColorIndex = 6 'gelb
        End If
    Next i
    
End Sub


' Beschreibung:
' Schreibt für jede Kogr die entsprechende der Einfärbung entsprechend der folgenden Regel
'   1. Wenn die KoGr nur ein Farbe hat, ist die Farbe in "Regel" Spalte geschrieben
'   2. Sonst, ist nichts geschrieben und die entsprechende Zelle ist als blau markiert

Sub RegelErzeugen(ws As Worksheet)

    Dim i As Long
    Dim j As Long
    Dim Zeilenanzahl As Long
    
    'PIVOT Tabelle Zeilenanzahl
    Let Zeilenanzahl = ws.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    
    'Rechnet Regel für jede GUID
    For i = 4 To Zeilenanzahl
        'If (Len(ws.Cells(i, 1)) = 32 And InStr(ws.Cells(i, 1), " ") = 0) Or ws.Cells(i, 1) = "(Leer)" Or ws.Cells(i, 1) = "Aus dem Motor" Then
        If (Len(ws.Cells(i, 1)) = 32 And InStr(ws.Cells(i, 1), " ") = 0) Then
            g = ws.Cells(i, 2)
            n = ws.Cells(i, 3)
            s = ws.Cells(i, 4)
            If g + n + s = g Then
                ws.Cells(i, 7) = "g"
            ElseIf g + n + s = n Then
                ws.Cells(i, 7) = "n"
            ElseIf g + n + s = s Then
                ws.Cells(i, 7) = "s"
            Else
                ws.Cells(i, 7).Interior.ColorIndex = 5 'color blue
            End If
        End If
    Next i
    
    'Kopf der Regel Spalten
    ws.Cells(2, 7) = "Regel"
    ws.Cells(2, 7).HorizontalAlignment = xlCenter
    ws.Cells(2, 7).VerticalAlignment = xlCenter
    ws.Cells(2, 7).Font.Bold = True
    
' 'Why do we need this additional ModulOrg data?
    ws.Cells(2, 8) = "ModulOrg"
    ws.Cells(2, 8).HorizontalAlignment = xlCenter
    ws.Cells(2, 8).VerticalAlignment = xlCenter
    ws.Cells(2, 8).Font.Bold = True

    'ModulOrg für Analyse in Erzeugnis Makro in Spalte 8 speichern
    For i = 5 To Zeilenanzahl
        If Len(ws.Cells(i, 1)) = 4 And IsNumeric(ws.Cells(i, 1).Value) = False Then
            ws.Cells(i, 8) = ws.Cells(i, 1)
        End If
    Next i

    For i = 5 To Zeilenanzahl
        j = 1
        If ws.Cells(i, 8) <> "" Then
            Do While ws.Cells(i + j, 8) = ""
                ws.Cells(i + j, 8) = ws.Cells(i, 8)
                j = j + 1
                If i + j > Zeilenanzahl Then
                    Exit Do
                End If
            Loop
        End If
    Next i
    
End Sub


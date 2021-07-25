Attribute VB_Name = "Allgemein_start"
Option Explicit

    Public der As String
    Public derCount As Integer


Sub go()

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With


    derCount = selectedDerivatCount
    If derCount = 1 Then
        '' if the derivat slicer exists then it's the one selected
        If ThisWorkbook.SlicerCaches.count > 0 Then
            der = ThisWorkbook.SlicerCaches("Datenschnitt_Derivat").VisibleSlicerItems(1).Name
        
        '' else if it doesn't exist, it's the first pivot item
        Else
            der = ThisWorkbook.Sheets("PIVOT").PivotTables("PivotTableMEGALISTE").PivotFields("Derivat").PivotItems(1)
        End If

        Call EinzelStart(der)

    ElseIf derCount > 1 Then
        Call GesamtStart
    End If

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With

End Sub

Sub EinzelStart(der As String)

    'Update März 2017:
    'Es ist nun möglich die Typ-Merkmale mit abzubilden, im Typschlüssel sind noch andere Merkmale falls gewünscht, dann muss der arr(x,x) verlängert werden
    'Sollten von einem Derivat mehrere Typschlüssel eingefügt werden, dann muss für die richtige aussage das auslesen des Typschlüssel zusätzlich gemacht werden
    'da momentan der eineindeutige über das Gültigkeitsatum gefunden wird, bei zwei Typschlüssel gäbe es dann aber zwei und dann muss nach einem eindeutigen Typschlüssel
    'gesucht werden. Hierzu möglicherweise Typschlüssel mit in Megaliste speichern - aber solange nur Berichtstyp passt es !!!!!!!!!!!!!!
    'Update März 2018: So kann man auch BZD oder BBD oder ZV Stände abbilden
    Dim shTyp As Worksheet
    Dim i As Integer, rw As Integer
    Dim j As Long
    Dim dataTyp() As Variant, arr() As Variant
           

    '' main function which calls steps one by one for the Einzeldarstellung
    Call EinzelPivot
    Call summaryTabelle
    Call wasserfallGraph
    Call formatTrep(der)
    Call formatPie(der)
    Call FBGraph(der)
    
'    ' 08.01.2019 Bigfix Aviv: keine FB Auswertung für Konfigprämissen Datensatz.
'    If InStr(der, "(KP)") = 0 Then
'        Call FBGraph(der)
'    End If
    
    Set shTyp = ThisWorkbook.Sheets("Typschl")
    dataTyp = shTyp.UsedRange
    
    rw = selectedDerivatCount
    ReDim arr(1 To rw, 1 To 5)
    
    rw = 0
    With ThisWorkbook.Sheets("PIVOT").PivotTables("PivotTableMEGALISTE")
        For i = 1 To .PivotFields("Derivat").PivotItems.count
            If .PivotFields("Derivat").PivotItems(i).Visible = True Then
                rw = rw + 1
                arr(rw, 1) = .PivotFields("Derivat").PivotItems(i).Name
                For j = 1 To UBound(dataTyp)
                    '' the SOP and Markt Segment are read from the Typschl table
                    If dataTyp(j, 2) = arr(rw, 1) And dataTyp(j, 7) <> vbNullString Then '' if gultigkeitdatum is filed
                        arr(rw, 2) = dataTyp(j, 7) 'Gültigkeit
                        arr(rw, 3) = dataTyp(j, 8) 'Motor
                        arr(rw, 4) = dataTyp(j, 10) 'Basis
                        'arr(rw, 5) = dataTyp(j, 12) 'E/EA/WA
                    End If
                Next j
            End If
        Next i
    End With
    
    '' set the button that generates the details for the gesamtdarstellung to false to prevent it from being clickable
    ThisWorkbook.Sheets("Home").detailMode.Value = False
    '' update the summary string above the Einzeldarstellung
    '' März 2017 Typmerkmale werden hier mit aufgenommen und werden als Filter für die Erstellung der Detailliste verwendet und sollte deswegen nciht angepasst werden, deswegen gesplittet in 2 Zellen
    If rw = 0 Then
        ThisWorkbook.Sheets("Home").Range("AO13").Value = "Einzeldarstellung " & SelectedSlicer
        ThisWorkbook.Sheets("Home").Range("A13").Value = "Einzeldarstellung " & SelectedSlicer
    Else
        ThisWorkbook.Sheets("Home").Range("AO13").Value = "Einzeldarstellung " & SelectedSlicer
        ThisWorkbook.Sheets("Home").Range("A13").Value = "Einzeldarstellung " & SelectedSlicer & " | Gültigkeitsdatum: " & arr(rw, 2) & " | Motor-Bezeichnung: " & arr(rw, 3) & " | Basisausführung: " & arr(rw, 4)
        ThisWorkbook.Sheets("Home").Range("AT21").Value = arr(rw, 5)
    End If
End Sub

Sub GesamtStart()
    
    
    '' main function which calls steps one by one for the Gesamtdarstellung
    Call GesamtPivot
    Call quelletab
    Call createHeatMap
    Call formatHeatMap
    Call addPieChart
    Call addArrow
    Call initializeArrow
    Call addSegmentLine
    Call scoring
    Call createScoringMap
    Call createGesamtGraphs
    
    ThisWorkbook.Sheets("Home").detailMode.Value = False
    ThisWorkbook.Sheets("Home").shapeButton.Value = False
    ThisWorkbook.Sheets("Home").Range("A41").Value = "Gesamtdarstellung " & SelectedSlicer
    '' lower the view to see directly the gesamtdarstellung chart
    ActiveWindow.ScrollRow = 41
End Sub

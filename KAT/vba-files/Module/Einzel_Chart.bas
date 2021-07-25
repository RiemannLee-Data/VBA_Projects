Attribute VB_Name = "Einzel_Chart"
Option Explicit

Sub summaryTabelle() 'put this in a array too

    Dim sp As Variant
    Dim sp2 As Variant
    Dim sp3 As Variant
    Dim sp4 As Variant
    Dim col As Integer, row As Integer
    Dim shPiv As Worksheet
    Dim tbl As ListObject
    Dim tbE As ListObject
    Dim tbEA As ListObject
    Dim tbWA As ListObject
    Dim i As Long, j As Integer
    Dim datapiv() As Variant
    Dim datapiv2() As Variant
    Dim datapiv3() As Variant
    Dim datapiv4() As Variant
    Dim matched As Boolean
    Dim Hilfe As String

    sp = Array("g", "gSA", "s", "sSA", "n", "nSA")
    sp2 = Array("EK_g", "EK_gSA", "EK_s", "EK_sSA", "EK_n", "EK_nSA", "ED_g", "ED_gSA", "ED_s", "ED_sSA", "ED_n", "ED_nSA")
    sp3 = Array("EAK_g", "EAK_gSA", "EAK_s", "EAK_sSA", "EAK_n", "EAK_nSA", "EAD_g", "EAD_gSA", "EAD_s", "EAD_sSA", "EAD_n", "EAD_nSA")
    sp4 = Array("WAK_g", "WAK_gSA", "WAK_s", "WAK_sSA", "WAK_n", "WAK_nSA", "WAD_g", "WAD_gSA", "WAD_s", "WAD_sSA", "WAD_n", "WAD_nSA")
    
    Set shPiv = ThisWorkbook.Sheets("PIVOT")
    Set tbl = ThisWorkbook.Sheets("Home").ListObjects("ZusTab")
    Set tbE = ThisWorkbook.Sheets("Home").ListObjects("Referenzauswertung")
    Set tbEA = ThisWorkbook.Sheets("Home").ListObjects("ReferenzauswertungEA")
    Set tbWA = ThisWorkbook.Sheets("Home").ListObjects("ReferenzauswertungWA")

    datapiv = shPiv.Range(shPiv.PivotTables("PivotTableMEGALISTE").TableRange1.Address)
    
    For i = 0 To UBound(sp)
        matched = False
        tbl.HeaderRowRange(i + 1) = sp(i)
        For j = 1 To UBound(datapiv, 2)
            If datapiv(2, j) = sp(i) & " Ergebnis" Then
                tbl.DataBodyRange(1, i + 1) = datapiv(UBound(datapiv), j)
                matched = True
                Exit For
            End If
        Next j
        If matched = False Then
            tbl.DataBodyRange(1, i + 1) = 0
        End If
    Next i
    
    tbl.DataBodyRange(1, 7) = "=SUM(ZusTab[@[g]:[nSA]])"
    
    col = tbl.DataBodyRange(1, 7).Column
    row = tbl.DataBodyRange(1, 7).row
    tbl.DataBodyRange(2, 1).FormulaR1C1 = "=R[-1]C/R" & row & "C" & col
    tbl.DataBodyRange(2, 1).AutoFill Destination:=tbl.DataBodyRange.Rows(2), Type:=xlFillValues
    tbl.DataBodyRange.Rows(2).Style = "Percent"
    
    If ThisWorkbook.Sheets("Home").nurBasis.Value = True Then
        tbl.DataBodyRange.Columns(2).ClearContents
        tbl.DataBodyRange.Columns(4).ClearContents
        tbl.DataBodyRange.Columns(6).ClearContents
    End If

    'Übertrag der ReferenzauswertungE
    datapiv2 = shPiv.Range(shPiv.PivotTables("PivotTableMEGALISTE").TableRange1.Address)

    For i = 0 To UBound(sp2)
        matched = False
        tbE.HeaderRowRange(i + 1) = sp2(i)
        For j = 1 To UBound(datapiv2, 2)
            'Hilfe = sp2(i) & " Ergebnis"
            If datapiv2(3, j) = sp2(i) & " Ergebnis" Then
                tbE.DataBodyRange(1, i + 1) = datapiv2(UBound(datapiv2), j)
                matched = True
                Exit For
            End If
        Next j
        If matched = False Then
            tbE.DataBodyRange(1, i + 1) = 0
        End If
    Next i

    If ThisWorkbook.Sheets("Home").nurBasis.Value = True Then
        tbE.DataBodyRange.Columns(2).ClearContents
        tbE.DataBodyRange.Columns(4).ClearContents
        tbE.DataBodyRange.Columns(6).ClearContents
        tbE.DataBodyRange.Columns(8).ClearContents
        tbE.DataBodyRange.Columns(10).ClearContents
        tbE.DataBodyRange.Columns(12).ClearContents
    End If
    
    'Übertrag der ReferenzauswertungEA
    datapiv3 = shPiv.Range(shPiv.PivotTables("PivotTableMEGALISTE").TableRange1.Address)
    
    For i = 0 To UBound(sp3)
        tbEA.DataBodyRange(1, i + 1) = 0
        matched = False
        tbEA.HeaderRowRange(i + 1) = sp3(i)
        For j = 1 To UBound(datapiv3, 2)
            If datapiv3(4, j) = sp3(i) & " Ergebnis" Then
                tbEA.DataBodyRange(1, i + 1) = ((tbEA.DataBodyRange(1, i + 1)) + (datapiv3(UBound(datapiv3), j)))
                matched = True
            End If
        Next j
        If matched = False Then
            tbEA.DataBodyRange(1, i + 1) = 0
        End If
    Next i
    
    If ThisWorkbook.Sheets("Home").nurBasis.Value = True Then
        tbEA.DataBodyRange.Columns(2).ClearContents
        tbEA.DataBodyRange.Columns(4).ClearContents
        tbEA.DataBodyRange.Columns(6).ClearContents
        tbEA.DataBodyRange.Columns(8).ClearContents
        tbEA.DataBodyRange.Columns(10).ClearContents
        tbEA.DataBodyRange.Columns(12).ClearContents
    End If
    
    'Übertrag der ReferenzauswertungWA
    datapiv4 = shPiv.Range(shPiv.PivotTables("PivotTableMEGALISTE").TableRange1.Address)
    
    For i = 0 To UBound(sp4)
        tbWA.DataBodyRange(1, i + 1) = 0
        matched = False
        tbWA.HeaderRowRange(i + 1) = sp4(i)
        For j = 1 To UBound(datapiv4, 2)
            If datapiv4(5, j) = sp4(i) Then
                tbWA.DataBodyRange(1, i + 1) = ((tbWA.DataBodyRange(1, i + 1)) + (datapiv4(UBound(datapiv4), j)))
                matched = True
                'Exit For
            End If
        Next j
        If matched = False Then
            tbWA.DataBodyRange(1, i + 1) = 0
        End If
    Next i
    
    If ThisWorkbook.Sheets("Home").nurBasis.Value = True Then
        tbEA.DataBodyRange.Columns(2).ClearContents
        tbEA.DataBodyRange.Columns(4).ClearContents
        tbEA.DataBodyRange.Columns(6).ClearContents
        tbEA.DataBodyRange.Columns(8).ClearContents
        tbEA.DataBodyRange.Columns(10).ClearContents
        tbEA.DataBodyRange.Columns(12).ClearContents
    End If

End Sub


Sub FBGraph(der As String)
'Sub FBGraph()
'    Dim der As String
'    der = "F66"
    
    Dim i As Integer, j As Integer, m As Integer, Fehler As Integer, max_row As Integer
    Dim tbl As ListObject, zusTbl As ListObject, pieTbl As ListObject
    Dim shPiv As Worksheet
    Dim datapiv() As Variant
    
    Set shPiv = ThisWorkbook.Sheets("PIVOT_FB")
    Set tbl = ThisWorkbook.Sheets("Home").ListObjects("TableFB")
    
  
    ' On Error GoTo Fehlerhandling
    With shPiv.PivotTables("PivotTableFB").PivotFields("Derivat")
        .ClearAllFilters
        .CurrentPage = der
    End With
    
    datapiv = shPiv.Range(shPiv.PivotTables("PivotTableFB").TableRange1.Address)

    ' die Fachbereich entspricht nicht immer die Ordnung geschrieben darunter
    ' so wir müssen zuerst die Zeile von datapiv ordnen?
    For i = 3 To 15
        If datapiv(i, 1) = "Gesamtergebnis" Then
            max_row = i
            Exit For
        End If
    Next i
    
    
    With tbl
        .Application.AutoCorrect.AutoFillFormulasInLists = False
        .DataBodyRange.ClearContents
        .DataBodyRange(1, 1) = "Gesamt"
        

        For i = 3 To (max_row - 1)
            .DataBodyRange(i - 1, 1) = datapiv(i, 1)
        Next i
        
        ' Here ist die Impliementierung, wie mann Daten in Worksheet "PIVOT_FB" in FB-Graphik dargestellen kann
        ' und hier haben wir den Code vereinfachen, der originale Code kann man in KAT5.3_inkl finden
        
        For j = 2 To 8

            ' schleifen in einer Reihe, falls wir "Gesamtergebnis" bekommen, dann ist das Schleife am Ende
            ' aber es gibt Fehler bei NA0, z.B.
            If datapiv(2, j) = "Gesamtergebnis" Then
                Exit For
            ElseIf datapiv(2, j) = "g" Then m = 2
            ElseIf datapiv(2, j) = "s" Then m = 4
            ElseIf datapiv(2, j) = "n" Then m = 6
            Else: GoTo NextIteration
            End If
            
            '' Wert ohne Sonderausstattung
            If ThisWorkbook.Sheets("Home").nurBasis.Value = True Then
                For i = 3 To UBound(datapiv)
                    ' schleifen in einer Kolumne, falls wir "Gesamtergebnis" bekommen, dann ist das Schleife am Ende
                    If datapiv(i, 1) = "Gesamtergebnis" Then
                        .DataBodyRange(1, m + 1) = datapiv(i, j)
                        Exit For
                    Else
                        .DataBodyRange(i - 1, m + 1) = datapiv(i, j)
                    End If
                Next i
            
            '' Wert mit Sonderausstattung
            Else
                For i = 3 To UBound(datapiv)
                    If datapiv(2, j + 1) = (datapiv(2, j) & "SA") Then
                        ' schleifen in einer Kolumne, falls wir "Gesamtergebnis" bekommen, dann ist das Schleife am Ende
                        If datapiv(i, 1) = "Gesamtergebnis" Then
                            .DataBodyRange(1, m + 1) = datapiv(i, j) + datapiv(i, j + 1)
                            Exit For
                        Else
                            .DataBodyRange(i - 1, m + 1) = datapiv(i, j) + datapiv(i, j + 1)
                        End If
                    
                    '' Es gibe Falls, wo gSA, nSA, sSA nicht bestehend ist, dann ist die Rechnung gleich wie ohne Sonderausstattung
                    Else
                        If datapiv(i, 1) = "Gesamtergebnis" Then
                            .DataBodyRange(1, m + 1) = datapiv(i, j)
                            Exit For
                        Else
                            .DataBodyRange(i - 1, m + 1) = datapiv(i, j)
                        End If
                    End If
                Next i

            End If
NextIteration:
        Next j
        
        On Error Resume Next
        
        
        For i = 8 To 1 Step -1
            If .DataBodyRange(i, 3) = 0 And .DataBodyRange(i, 5) = 0 And .DataBodyRange(i, 7) = 0 Then
                .ListRows(i).Delete
            Else
                Exit For
            End If
        Next i

        .DataBodyRange(3, 2).FormulaR1C1 = "=R[-1]C[1]"
        .DataBodyRange(4, 2).FormulaR1C1 = "=R[-1]C[1]+R[-1]C"
        .DataBodyRange(4, 2).AutoFill Destination:=.DataBodyRange.Columns(2).Rows("4:" & i), Type:=xlFillDefault

        .DataBodyRange(2, 4).FormulaR1C1 = "=R[-1]C[-1]-RC[-1]"
        .DataBodyRange(3, 4).FormulaR1C1 = "=R[-1]C[+1]-RC[-1]+R[-1]C"
        .DataBodyRange(3, 4).AutoFill Destination:=.DataBodyRange.Columns(4).Rows("3:" & i), Type:=xlFillDefault

        .DataBodyRange(2, 6).FormulaR1C1 = "=R[-1]C[-1]-RC[-1]"
        .DataBodyRange(3, 6).FormulaR1C1 = "=R[-1]C[+1]-RC[-1]+R[-1]C"
        .DataBodyRange(3, 6).AutoFill Destination:=.DataBodyRange.Columns(6).Rows("3:" & i), Type:=xlFillDefault
        
        On Error GoTo 0
    End With

    ThisWorkbook.Sheets("Home").ChartObjects("trepFB").Chart.FullSeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 255, 0) '' green
        .Transparency = 0
        .Solid
    End With
    
    ActiveChart.FullSeriesCollection(4).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0) '' yellow
        .Transparency = 0
        .Solid
    End With
    
    ActiveChart.FullSeriesCollection(6).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0) '' red
        .Transparency = 0
        .Solid
    End With

    With ThisWorkbook.Sheets("Home").ChartObjects("trepFB").Chart
        .FullSeriesCollection(1).Format.Fill.Visible = msoFalse
        .FullSeriesCollection(1).HasDataLabels = False
        .FullSeriesCollection(3).Format.Fill.Visible = msoFalse
        .FullSeriesCollection(3).HasDataLabels = False
        .FullSeriesCollection(5).Format.Fill.Visible = msoFalse
        .FullSeriesCollection(5).HasDataLabels = False
        .FullSeriesCollection(6).Points(1).ApplyDataLabels
        .SetElement (msoElementLegendRight)
        .SetElement (msoElementChartTitleAboveChart)
        .SetElement (msoElementPrimaryCategoryAxisShow)
        .ChartTitle.Text = der
        .ChartTitle.Format.TextFrame2.TextRange.Characters.Text = der

        With .ChartTitle.Format.TextFrame2.TextRange.Characters.ParagraphFormat
            .TextDirection = msoTextDirectionLeftToRight
            .Alignment = msoAlignCenter
        End With

        With .ChartTitle.Format.TextFrame2.TextRange.Characters.Font
            .BaselineOffset = 0
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(89, 89, 89)
            .Fill.Transparency = 0
            .Fill.Solid
            .size = 14
            .Italic = msoFalse
            .Kerning = 12
            .UnderlineStyle = msoNoUnderline
            .Spacing = 0
            .Strike = msoNoStrike
            .NameComplexScript = "BMWType V2 Light"
            .NameFarEast = "BMWType V2 Light"
            .Name = "BMWType V2 Light"
            .Bold = msoTrue
        End With
    End With
    
    ThisWorkbook.Sheets("Home").Cells(1, 1).Select
   ' Set the visibility of the sheet "PIVOT_FB" with this variable here
   ' Seems better to make it visible for comparison when test the data
'    shPiv.Visible = xlSheetHidden
    

End Sub


Sub wasserfallGraph()

    Dim shPiv As Worksheet
    Dim tbl As ListObject, zusTbl As ListObject, pieTbl As ListObject
    Dim i As Integer, j As Integer, topfive As Integer
    Dim gCol As Integer, gSACol As Integer, sCol As Integer, sSACol As Integer
    Dim datapiv() As Variant
    
    ' to do: possibility to store everything in an array before pasting the result directly in a table list object
    Set tbl = ThisWorkbook.Sheets("Home").ListObjects("wasserfallTab")
    Set zusTbl = ThisWorkbook.Sheets("Home").ListObjects("ZusTab")
    Set pieTbl = ThisWorkbook.Sheets("Home").ListObjects("pieTab")
    Set shPiv = ThisWorkbook.Sheets("PIVOT")
    
    shPiv.PivotTables("PivotTableMEGALISTE").PivotFields("Kommunalität").ShowDetail = False
    
    With tbl
        Application.AutoCorrect.AutoFillFormulasInLists = False
        .DataBodyRange.ClearContents
        .DataBodyRange(1, 1) = "Gesamt"
        .DataBodyRange(7, 1) = "Rest"
        
        '' fill in Gesamt line according to the ZusTab
        .DataBodyRange(1, 3) = zusTbl.DataBodyRange(1, 1) + zusTbl.DataBodyRange(1, 2)
        pieTbl.DataBodyRange(1, 2) = .DataBodyRange(1, 3)
        .DataBodyRange(1, 5) = zusTbl.DataBodyRange(1, 3) + zusTbl.DataBodyRange(1, 4)
        pieTbl.DataBodyRange(2, 2) = .DataBodyRange(1, 5)
        .DataBodyRange(1, 6) = zusTbl.DataBodyRange(1, 5) + zusTbl.DataBodyRange(1, 6)
        pieTbl.DataBodyRange(3, 2) = .DataBodyRange(1, 6)
        
        datapiv = shPiv.Range(shPiv.PivotTables("PivotTableMEGALISTE").TableRange1.Address)
        
        gCol = 0: gSACol = 0: sCol = 0: sSACol = 0
        For j = 1 To UBound(datapiv, 2)
            If datapiv(2, j) = "g" Then gCol = j
            If datapiv(2, j) = "gSA" Then gSACol = j
            If datapiv(2, j) = "s" Then sCol = j
            If datapiv(2, j) = "sSA" Then sSACol = j
        Next j
                
        topfive = 1
        For i = 6 To UBound(datapiv) 'Aviv 14.11.2018 Bugfix war 7
            If datapiv(i, 1) = "Gesamtergebnis" Then Exit For
            If topfive > 5 Then Exit For
            
            If datapiv(i, 1) <> "(Leer)" Then

                .DataBodyRange(topfive + 1, 1) = datapiv(i, 1)
                If gCol <> 0 Then
                    .DataBodyRange(topfive + 1, 3) = datapiv(i, gCol)
                End If
                
                If gSACol <> 0 And ThisWorkbook.Sheets("Home").nurBasis.Value = False Then
                    .DataBodyRange(topfive + 1, 3) = .DataBodyRange(topfive + 1, 3) + datapiv(i, gSACol)
                End If
                
                If sCol <> 0 Then
                    .DataBodyRange(topfive + 1, 5) = datapiv(i, sCol)
                End If
                
                If sSACol <> 0 And ThisWorkbook.Sheets("Home").nurBasis.Value = False Then
                    .DataBodyRange(topfive + 1, 5) = .DataBodyRange(topfive + 1, 5) + datapiv(i, sSACol)
                End If
                
                topfive = topfive + 1
            End If
        Next i
        
        .DataBodyRange(7, 3).FormulaR1C1 = "=R[-6]C-SUM(R[-5]C:R[-1]C)"
        .DataBodyRange(7, 5).FormulaR1C1 = .DataBodyRange(7, 3).FormulaR1C1
        
        ' fill in column value for the blank in between GT and ST value to allow a cascade display
        .DataBodyRange(3, 2).FormulaR1C1 = "=R[-1]C[1]"
        .DataBodyRange(4, 2).FormulaR1C1 = "=R[-1]C[1]+R[-1]C"
        .DataBodyRange(4, 2).AutoFill Destination:=.DataBodyRange.Columns(2).Rows("4:7"), Type:=xlFillDefault
        .DataBodyRange(2, 4).FormulaR1C1 = "=R[-1]C[-1]-RC[-1]"
        .DataBodyRange(3, 4).FormulaR1C1 = "=R[-1]C[+1]-RC[-1]+R[-1]C"
        .DataBodyRange(3, 4).AutoFill Destination:=.DataBodyRange.Columns(4).Rows("3:7"), Type:=xlFillDefault
        
    End With
    
    shPiv.PivotTables("PivotTableMEGALISTE").PivotFields("Kommunalität").ShowDetail = True
    
End Sub

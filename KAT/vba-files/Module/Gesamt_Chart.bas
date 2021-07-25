Attribute VB_Name = "Gesamt_Chart"
Option Explicit



Sub createHeatMap()

    '' this recreates the chart from scratch
    Dim ws As Worksheet
    Dim objChrt As ChartObject
    Dim s As Series
    Dim lngIndex As Integer
    Dim Derivat As String
    Dim size As Integer, varSize
    Dim Shp As Shape
    
    Set ws = ThisWorkbook.Sheets("Home")
    With ws
        '' first we delete the existing charts to be sure to take a fresh start and not take over any weird elements from the old one
        On Error Resume Next
        .ChartObjects("HeatMap").Delete
        On Error GoTo 0
        
        '' this parameter allows the user to change the charts size according to his needs
        varSize = ThisWorkbook.Sheets("Home").Range("B42").Value
        If VarType(varSize) = 5 And varSize >= 0 Then '5 = double
            size = varSize
        Else
            size = 0
        End If
        
        Set objChrt = .ChartObjects.add(180, 601, 600 + 100 * size, 400 + 70 * size) '' the width and height change witht the size
        
        With objChrt.Chart
            .Parent.Name = "HeatMap" '' the chart object
            
            Set s = .SeriesCollection.NewSeries()
            
            With s
                .ChartType = xlXYScatter
                .Name = "Gesamtdarstellung" '' give it a title
                .XValues = ws.ListObjects("quelleTab").DataBodyRange.Columns(2) '' assign the SOP as x
                .Values = ws.ListObjects("quelleTab").DataBodyRange.Columns(4) '' assign "werte" as y
    
                For lngIndex = 1 To .Points.count
                    Derivat = ThisWorkbook.Sheets("Home").ListObjects("quelleTab").DataBodyRange(lngIndex, 1)
                    .Points(lngIndex).HasDataLabel = True
                    .Points(lngIndex).DataLabel.Text = Derivat '' give each point a individual name
                Next lngIndex
            End With
            '' rotate the labels in x-axis for better visualization
            .Axes(xlCategory).TickLabels.Orientation = 45
            
        End With
    End With
End Sub



'' the order UKL1, UKL2, KKL, MKL, GKL is set by the sub sortquelle, in module Gesamt_Quelle. To add a new markt segment, do it there
Sub addSegmentLine()

    Dim Ynode1 As Double, Ymin As Double, Ymax As Double
    Dim Ytop As Double, Xleft As Double, Yheight As Double, Xwidth As Double
    Dim tbl As ListObject
    Dim mkt As String
    Dim i As Integer
    
    Set tbl = ThisWorkbook.Sheets("Home").ListObjects("quelleTab")
    
    ' name and format shape as arrowhead
    With ThisWorkbook.Sheets("Home").ChartObjects("HeatMap").Chart
        '' allows to measure the chart area
        Ytop = .PlotArea.InsideTop
        Yheight = .PlotArea.InsideHeight
        Xleft = .PlotArea.InsideLeft
        Xwidth = .PlotArea.InsideWidth
        Ymin = .Axes(2).MinimumScale
        Ymax = .Axes(2).MaximumScale
        
        mkt = vbNullString
        
        '' loop through markt segment : for each new market segment, trace a lign under it's first point
        For i = 1 To tbl.DataBodyRange.Rows.count
            If tbl.DataBodyRange(i, 3) <> mkt Then
                mkt = tbl.DataBodyRange(i, 3)
                Ynode1 = Ytop + 10 + (Ymax - tbl.DataBodyRange(i, 4)) * Yheight / (Ymax - Ymin)
                
                With .Shapes.AddLine(0, Ynode1, 2 * Xleft + Xwidth, Ynode1)
                    .Name = "SegmentLine" & mkt
                    .Line.ForeColor.RGB = RGB(45, 45, 45)
                End With
                
                With .Shapes.AddTextbox(msoTextOrientationHorizontal, 0, Ynode1 - 17, 60, 15)
                    .TextFrame.Characters.Text = mkt '' add a text area mentionning the Markt Segment
                    .Name = "MarktSegment" & mkt
                End With
                
            End If
        Next i

    End With
End Sub



' here we kind of copy the function "FBGraph" in Einzel_Chart
Sub createGesamtGraphs()

    Dim Derivat() As String, der As Variant, Result As String
    Dim i As Integer, j As Integer, m As Integer, lngIndex As Integer, max_row As Integer
    Dim tbl_quelle As ListObject, tbl_gesamt_fb As ListObject, tbl_gesamt_pie As ListObject
    Dim datapiv() As Variant
    Dim shPiv As Worksheet

    Set shPiv = ThisWorkbook.Sheets("PIVOT_FB")
    Set tbl_quelle = ThisWorkbook.Sheets("Home").ListObjects("quelleTab")
    Set tbl_gesamt_fb = ThisWorkbook.Sheets("Home").ListObjects("GesamtTableFB")
    Set tbl_gesamt_pie = ThisWorkbook.Sheets("Home").ListObjects("GesamtPieTab")
    
    ' vba doesn't allow for appending items to an array
    ' thus we need redimension to complete it
    For lngIndex = 1 To tbl_quelle.DataBodyRange.Rows.count
        ReDim Preserve Derivat(lngIndex - 1)
        Derivat(lngIndex - 1) = tbl_quelle.DataBodyRange(lngIndex, 1)
    Next lngIndex
    
    Result = UCase(Join(Derivat, "|"))
    
    ' here we need to enable multiple choice of derivat
    With shPiv.PivotTables("PivotTableFB").PivotFields("Derivat")
        .ClearAllFilters
        .EnableMultiplePageItems = True
        .PivotItems(1).Visible = True
        
        ' the code here should be optimized later
        For i = 2 To .PivotItems.count
            If .PivotItems(i).Visible Then .PivotItems(i).Visible = False
        Next i
        
        For Each der In Derivat
            .PivotItems(der).Visible = True
        Next der
        
        If InStr(Result, UCase(.PivotItems(1))) = 0 Then .PivotItems(1).Visible = False
        
        If Err.Number <> 0 Then
            .ClearAllFilters
            MsgBox Title:="No Items Found", Prompt:="None of the desired items was found in the Pivot, so I have cleared the filter"
        End If
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
    
    
    With tbl_gesamt_fb
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
        
        ' divide the number of derivat to get the average data
        '.DataBodyRange.Value = .DataBodyRange.Value / (lngIndex - 1)
        
        For m = 1 To 3
            tbl_gesamt_pie.DataBodyRange(m, 2) = .DataBodyRange(1, 1 + 2 * m)
        Next m
        
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

    
    ThisWorkbook.Sheets("Home").ChartObjects("trepGesamt").Chart.FullSeriesCollection(2).Select
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
    
'''''''''''''''''''''''' formating of the waterfall chart ''''''''''''''''''''''''
    With ThisWorkbook.Sheets("Home").ChartObjects("trepGesamt").Chart
        .FullSeriesCollection(1).Format.Fill.Visible = msoFalse
        .FullSeriesCollection(1).HasDataLabels = False
        .FullSeriesCollection(2).HasDataLabels = True
        .FullSeriesCollection(3).Format.Fill.Visible = msoFalse
        .FullSeriesCollection(3).HasDataLabels = False
        .FullSeriesCollection(4).HasDataLabels = True
        .FullSeriesCollection(5).Format.Fill.Visible = msoFalse
        .FullSeriesCollection(5).HasDataLabels = False
        .FullSeriesCollection(6).HasDataLabels = True
        .FullSeriesCollection(6).Points(1).ApplyDataLabels
        .SetElement (msoElementLegendRight)
        .SetElement (msoElementChartTitleAboveChart)
        .SetElement (msoElementPrimaryCategoryAxisShow)
        .ChartTitle.Text = "Gesamt TrepFB:  " + Join(Derivat, ",")
        '.ChartTitle.Format.TextFrame2.TextRange.Characters.Text = Derivat

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
    
'''''''''''''''''''''''' formating of the pie chart ''''''''''''''''''''''''

    With ThisWorkbook.Sheets("Home").ChartObjects("pieDiaGesamt")
        .ShapeRange.Line.Visible = msoTrue
        
        '' msoTrue is a Long with a value of -1
        '' True is a Boolean, which is stored as a Long with a value of -1
        '' So in the Immediate window, "?msoTrue = True" returns "Wahr"

        With .Chart
            .SetElement (msoElementChartTitleAboveChart)
            .SetElement (msoElementLegendLeft)
            .ChartTitle.Text = "Gesamt Pie Chart:  " + Join(Derivat, ",")
            '.ChartTitle.Format.TextFrame2.TextRange.Characters.Text = der
            
            With .ChartTitle.Format.TextFrame2.TextRange.Characters.ParagraphFormat
                .TextDirection = msoTextDirectionLeftToRight
                .Alignment = msoAlignCenter
            End With
            
            With .ChartTitle.Format.TextFrame2.TextRange.Characters.Font
                .BaselineOffset = 0
                .Fill.Visible = msoTrue
                .Fill.ForeColor.RGB = RGB(89, 89, 89)
                .Fill.Transparency = 0.01
                .Fill.Solid
                .size = 14
                .Italic = msoFalse
                .Kerning = 12
                .Name = "+mn-lt"
                .UnderlineStyle = msoNoUnderline
                .Spacing = 0
                .Strike = msoNoStrike
                .NameComplexScript = "BMWType V2 Light"
                .NameFarEast = "BMWType V2 Light"
                .Name = "BMWType V2 Light"
                .Bold = msoTrue
            End With
            
            With .FullSeriesCollection(1)
                .ApplyDataLabels
                .HasLeaderLines = True
                .DataLabels.ShowPercentage = True
                .DataLabels.ShowCategoryName = False
                .DataLabels.ShowValue = False
                .DataLabels.ShowSeriesName = False
                .DataLabels.ShowRange = False
                .DataLabels.Separator = "; "
                .DataLabels.Position = xlLabelPositionBestFit
                .Points(1).Format.Fill.ForeColor.RGB = RGB(0, 255, 0) '' green
                .Points(2).Format.Fill.ForeColor.RGB = RGB(255, 255, 0) '' yellow
                .Points(3).Format.Fill.ForeColor.RGB = RGB(255, 0, 0) '' red
            End With
    
        End With
    End With

    ThisWorkbook.Sheets("Home").Cells(42, 26).Select
   ' Set the visibility of the sheet "PIVOT_FB" with this variable here
   ' Seems better to make it visible for comparison when test the data
'    shPiv.Visible = xlSheetHidden
        
End Sub

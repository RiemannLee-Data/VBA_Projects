Attribute VB_Name = "Allgemein_PieChart"
Option Explicit


Sub saveAsPicture()

    Dim objChrt As ChartObject
    Dim myfilename As String
    Dim pfad As String
    Dim der As String

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    '' initialize objects, get derivat name from the chart
    Set objChrt = ThisWorkbook.Sheets("Home").ChartObjects("pieDia")
    der = objChrt.Chart.ChartTitle.Caption
    myfilename = der & ".png"
    
    '' delete pie chart with the name "der" in the heatmap chart diagramm folder
    Call deletePie(der)
    
    '' reformat the pie chart diagramm so that its clean and square
    objChrt.Chart.SetElement (msoElementChartTitleNone)
    objChrt.Chart.SetElement (msoElementLegendNone)
    objChrt.Chart.SetElement (msoElementDataLabelNone)
    objChrt.ShapeRange.Line.Visible = msoFalse
    
    '' save the pie chart diagram as image
    pfad = ThisWorkbook.Path & "\KAT_Vorlage\06_Heatmap_Chart_Diagramm\" & myfilename
    objChrt.Chart.Export fileName:=pfad, Filtername:="png"
    
    '' reformat the pie chart diagram
    Call formatPie(der)
    
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    
End Sub



Sub addPieChart()
    Dim sh As Worksheet
    Dim ch As Chart
    Dim der As String
    Dim str As String, pfad As String
    Dim tbl As ListObject
    Dim i As Integer
    
    Dim Xnode As Double, Ynode As Double
    Dim Xmin As Double, Xmax As Double, Ymin As Double, Ymax As Double
    Dim Xleft As Double, Ytop As Double, Xwidth As Double, Yheight As Double
    
    '' initialize variable
    Set sh = ThisWorkbook.Sheets("Home")
    Set ch = sh.ChartObjects("HeatMap").Chart
    Set tbl = ThisWorkbook.Sheets("Home").ListObjects("quelleTab")
    
    pfad = ThisWorkbook.Path & "\KAT_Vorlage\06_Heatmap_Chart_Diagramm\"
    
    With ch
        '' get the gesamtdarstellung measurements
        Xleft = .PlotArea.InsideLeft
        Xwidth = .PlotArea.InsideWidth
        Ytop = .PlotArea.InsideTop
        Yheight = .PlotArea.InsideHeight
        Xmin = .Axes(1).MinimumScale
        Xmax = .Axes(1).MaximumScale
        Ymin = .Axes(2).MinimumScale
        Ymax = .Axes(2).MaximumScale
        
        '' loop through all derivat and add a pie chart according to the poitn placement
        For i = 1 To tbl.DataBodyRange.Rows.count
            der = .FullSeriesCollection(1).Points(i).DataLabel.Text
            If der <> tbl.DataBodyRange(i, 1) Then
                MsgBox "OMG" & der & tbl.DataBodyRange(i, 1)
            End If
            
            If Dir(pfad & der & ".png") <> vbNullString Then
                Xnode = Xleft - 60 + (tbl.DataBodyRange(i, 2) - Xmin) * Xwidth / (Xmax - Xmin)
                Ynode = Ytop - 60 + (Ymax - tbl.DataBodyRange(i, 4)) * Yheight / (Ymax - Ymin)
                With .Shapes.AddPicture(pfad & der & ".png", msoCTrue, msoCTrue, Xnode, Ynode, 55, 55)
                    .Name = "pie_" & der
                    .AutoShapeType = msoShapeOval
                End With
            End If
        Next i
        
    End With
    
    '' set the button which hides pie to false
    ThisWorkbook.Sheets("Home").shapeButton.Value = False
    
End Sub



Sub hidePie()

    Dim Shp As Shape
    
    '' loop through shapes
    '' the shape button is the one which hides or show pie charts
    '' hide or show on the name criterium : pie_
    With ThisWorkbook.Sheets("Home")
        If .shapeButton.Value = True Then
            For Each Shp In .ChartObjects("HeatMap").Chart.Shapes
                If Left(Shp.Name, 3) = "pie" Then
                    Shp.Visible = False
                End If
            Next Shp
        ElseIf .shapeButton.Value = False Then
            For Each Shp In .ChartObjects("HeatMap").Chart.Shapes
                If Left(Shp.Name, 3) = "pie" Then
                    Shp.Visible = True
'                    Debug.Print Shp.Name
'                    ThisWorkbook.Sheets("Home").Shapes(Array(Shp, "HeatMap")).Group
                End If
            Next Shp
        End If
    End With
    
    
    
End Sub



Sub deletePie(der As String)

    Dim pfad As String
    'delete pie image of derivat
    pfad = ThisWorkbook.Path & "\KAT_Vorlage\06_Heatmap_Chart_Diagramm\" & der & ".png"
    On Error Resume Next
    Kill pfad
    On Error GoTo 0
    
End Sub


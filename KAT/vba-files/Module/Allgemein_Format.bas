Attribute VB_Name = "Allgemein_Format"
Option Explicit

Sub formatPie(der As String)
    
    ' This sub is responsible for the formating of the <<pie>> chart
    
    With ThisWorkbook.Sheets("Home").ChartObjects("pieDia")
        .ShapeRange.Line.Visible = msoTrue
        
        '' msoTrue is a Long with a value of -1
        '' True is a Boolean, which is stored as a Long with a value of -1
        '' So in the Immediate window, "?msoTrue = True" returns "Wahr"

        With .Chart
            .SetElement (msoElementChartTitleAboveChart)
            .SetElement (msoElementLegendLeft)
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
End Sub


' This sub is responsible for the formating of the <<waterfall>> chart
Sub formatTrep(der As String)
          
    With ThisWorkbook.Sheets("Home").ChartObjects("trepDia").Chart
    
        .FullSeriesCollection(1).Format.Fill.Visible = msoFalse
        .FullSeriesCollection(1).HasDataLabels = False
        .FullSeriesCollection(2).HasDataLabels = True
        .FullSeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(0, 255, 0)   '' green
        .FullSeriesCollection(3).Format.Fill.Visible = msoFalse
        .FullSeriesCollection(3).HasDataLabels = False
        .FullSeriesCollection(4).HasDataLabels = True
        .FullSeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(255, 255, 0) '' yellow
        .FullSeriesCollection(5).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)   '' red
        .FullSeriesCollection(5).HasDataLabels = False
        .FullSeriesCollection(5).Points(1).ApplyDataLabels
        
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
End Sub


Sub formatHeatMap()
    
    ' This sub is responsible for the formating of the <<HeatMap>> chart
    With ThisWorkbook.Sheets("Home").ChartObjects("HeatMap").Chart
        .SetElement (msoElementLegendNone)
        .SetElement (msoElementChartTitleNone)
        .SetElement (msoElementPrimaryValueGridLinesNone)
        .SetElement (msoElementPrimaryValueAxisNone)
        .Axes(xlCategory).MajorUnit = 90
        .Axes(xlCategory).TickLabels.NumberFormat = "MM.YYYY"
        
        With .FullSeriesCollection(1)
            .MarkerStyle = xlCircle
            .MarkerSize = 12
            
            With .Format.ThreeD
                .BevelTopType = msoBevelCircle
                .BevelTopInset = 6
                .BevelTopDepth = 6
            End With
            
            .Format.Line.Visible = msoFalse
            .DataLabels.Position = xlLabelPositionRight
            
            With .DataLabels.Format.TextFrame2.TextRange.Font
                .size = 14
                .NameComplexScript = "BMWType V2 Light"
                .Name = "BMWType V2 Light"
                .Bold = msoTrue
            End With
        End With
    End With
End Sub

Sub formatScoring()

    ' This sub is responsible for the formating of the <<scoring>> chart
    
    With ThisWorkbook.Sheets("Home").ChartObjects("ScoringDia").Chart
        .SetElement (msoElementLegendNone) ' deletes legend
        .SetElement (msoElementChartTitleAboveChart) ' adds title
        .SetElement (msoElementPrimaryCategoryAxisShow) ' add x_axis
        .ChartTitle.Font.size = 7
        .ChartTitle.Font.Bold = False
        .ChartTitle.Text = Left(SelectedSlicer, 252) & "..."
        .SetElement (msoElementPrimaryValueGridLinesNone)
        .Axes(xlCategory, xlPrimary).HasTitle = True 'x_axis title
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "common und synergy parts in derivat _i / total number of parts in derivat_i" 'Imitation Capacity
        .Axes(xlValue, xlPrimary).HasTitle = True 'y_axis title
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Number of carry-over parts from derivat_i / number of new parts in derivat_i" 'Innovation Capacity
        
        With .FullSeriesCollection(1) ' formats the data points the same as in the Gersamtdarstellung
            .MarkerStyle = xlCircle
            .MarkerSize = 12
            
            With .Format.ThreeD
                .BevelTopType = msoBevelCircle
                .BevelTopInset = 6
                .BevelTopDepth = 6
            End With
            
            .Format.Line.Visible = msoFalse
            .DataLabels.Position = xlLabelPositionRight
            
            With .DataLabels.Format.TextFrame2.TextRange.Font
                .size = 14
                .NameComplexScript = "BMWType V2 Light"
                .Name = "BMWType V2 Light"
                .Bold = msoTrue
            End With
            
        End With
        
        With .PlotArea.Format.Fill 'add color gradation in the background
            .TwoColorGradient msoGradientDiagonalDown, 1
            .ForeColor.RGB = RGB(160, 250, 170)
            .BackColor.RGB = RGB(240, 100, 100)
            .GradientStops.Insert RGB(250, 250, 150), 0.5
        End With
        
    End With
    
End Sub

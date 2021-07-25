Attribute VB_Name = "Gesamt_Arrow"
Option Explicit

Sub addArrow()
    
    'calculation variables
    Dim der As String, fzg As String
    Dim tbl As ListObject
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim gesamt As Integer, anzahlTeil As Integer
    Dim shPiv As Worksheet, shTyp As Worksheet
    Dim prozent() As Variant, lineweight() As Variant, linetransparancy() As Variant
    Dim datapiv() As Variant, dataQu() As Variant, dataTyp() As Variant
    Dim found As Boolean
    
    'drawing variables
    Dim Xnode1 As Double, Ynode1 As Double
    Dim Xnode2 As Double, Ynode2 As Double
    Dim Xmin As Double, Xmax As Double
    Dim Ymin As Double, Ymax As Double
    Dim Xleft As Double, Ytop As Double
    Dim Xwidth As Double, Yheight As Double
    
    
    Set tbl = ThisWorkbook.Sheets("Home").ListObjects("quelleTab")
    Set shPiv = ThisWorkbook.Sheets("PIVOT")
    Set shTyp = ThisWorkbook.Sheets("Typschl")
    dataTyp = shTyp.UsedRange
    datapiv = shPiv.Range(shPiv.PivotTables("PivotTableMEGALISTE").TableRange1.Address)
    dataQu = tbl.DataBodyRange
    
    prozent = Array(0.8, 0.6, 0.5, 0.4, 0.3, 0.2, 0.1, 0.05, 0)
    lineweight = Array(17, 15, 13, 11, 8, 6, 4, 3, 0)
    linetransparancy = Array(0.1, 0.15, 0.25, 0.3, 0.35, 0.4, 0.45, 0.5, 0.55)
    
    With ThisWorkbook.Sheets("Home").ChartObjects("HeatMap").Chart
        Xleft = .PlotArea.InsideLeft
        Xwidth = .PlotArea.InsideWidth
        Ytop = .PlotArea.InsideTop
        Yheight = .PlotArea.InsideHeight
        Xmin = .Axes(1).MinimumScale
        Xmax = .Axes(1).MaximumScale
        Ymin = .Axes(2).MinimumScale
        Ymax = .Axes(2).MaximumScale
        
        '' we loop twice in the derivat names
        '' the first loop gives the derivat name and the second the Fahrzeugsbezugseil
        '' in other words the first loop sets the target of the arrow and the second loop, the origin of the arrow
        For i = 1 To UBound(dataQu)
            der = dataQu(i, 1)
            gesamt = 0
            For j = 1 To UBound(dataTyp, 1)
                If dataTyp(j, 2) = der And dataTyp(j, 7) <> vbNullString Then
                    '' this gesamt value is very important. it sets the total number of parts in a derivat, regardless of filters
                    '' it will allow us to compute percents for each arrow
                    gesamt = dataTyp(j, 6)
                End If
            Next j
            If gesamt = 0 Or gesamt = Empty Then
                MsgBox "Gesamt value is empty or null for derivat " & der
            Else
                For j = 1 To UBound(dataQu)
                    If i <> j Then
                        fzg = dataQu(j, 1)
                        anzahlTeil = 0
                        found = False
                        For k = 1 To UBound(datapiv)
                            If datapiv(k, 1) = fzg Then
                                For l = 1 To UBound(datapiv, 2)
                                    If datapiv(2, l) = der Then
                                        anzahlTeil = datapiv(k, l) '' gets the value of parts from fzg to der
                                        found = True
                                        Exit For
                                    End If
                                Next l
                                If found = True Then Exit For
                            End If
                        Next k
                        If anzahlTeil > 0 And gesamt > 0 Then
                            '' compute the arrow's coordinates on chart
                            Xnode1 = Xleft + (dataQu(j, 2) - Xmin) * Xwidth / (Xmax - Xmin)
                            Ynode1 = Ytop + (Ymax - dataQu(j, 4)) * Yheight / (Ymax - Ymin)
                            Xnode2 = Xleft + (dataQu(i, 2) - Xmin) * Xwidth / (Xmax - Xmin)
                            Ynode2 = Ytop + (Ymax - dataQu(i, 4)) * Yheight / (Ymax - Ymin)
        
                            '' name and format shape as arrow
                            With .Shapes.AddLine(Xnode1, Ynode1, Xnode2, Ynode2)
                                .Name = "Arrow" & fzg & "-" & der
                                With .Line
                                    .EndArrowheadStyle = msoArrowheadTriangle
                                    .ForeColor.RGB = 5921370 'grey
                                    For k = 0 To UBound(prozent)
                                        If (anzahlTeil / gesamt) > prozent(k) Then
                                            '' gives line weight and transparency according to table set up early in sub-module
                                            .weight = lineweight(k)
                                            .Transparency = linetransparancy(k)
                                            Exit For
                                        End If
                                    Next k
                                End With
                            End With
                        End If
                    End If
                Next j
            End If
        Next i
    End With
End Sub


'' this sub-module assigns the 'highlightArrow' macro to the arrows on the worksheet (but not the ones on the gesamtdarstellung chart)
'' we pass the line weight and the color as arguments, they will control the arrows in the gesamtdarstellung chart
Sub initializeArrow()

    Dim Shp As Shape

    For Each Shp In ThisWorkbook.Sheets("Home").Shapes
        If Left(Shp.Name, 5) = "Arrow" Then
            Shp.Line.ForeColor.RGB = 5921370
            Shp.OnAction = "'highlightArrow """ & Shp.Line.weight & """,""" & Shp.Line.ForeColor.RGB & """'"
        End If
    Next Shp
    
End Sub



Sub highlightArrow(wght As Integer, clr As Long)

    'blue : RGB(150, 200, 255) = 16763030
    'grey : RGB(90, 90, 90) = 5921370
    Dim color As String
    Dim Shp As Shape
    '' this is a cyclic function
    '' normal -> hide -> highlight
    '' normal means grey in the legend and on the chart for each arrows with the same line weight
    '' brown in the legend means hide on the chart
    '' blue is a highlight on both the legend and the chart
    For Each Shp In ThisWorkbook.Sheets("Home").Shapes
        If Left(Shp.Name, 5) = "Arrow" Then
            If Shp.Line.weight = wght Then
                If Shp.Line.ForeColor.RGB = 232132 Then 'if whatever then blue
                    Shp.Line.ForeColor.RGB = 16763030
                    color = "blue"
                'Test new red
                ElseIf Shp.Line.ForeColor.RGB = 16763030 Then
                    Shp.Line.ForeColor.RGB = RGB(255, 0, 0) 'else red
                    color = "red"
                'Test new green
                ElseIf Shp.Line.ForeColor.RGB = RGB(255, 0, 0) Then
                    Shp.Line.ForeColor.RGB = RGB(84, 130, 53) 'else red
                    color = "green"
                ElseIf Shp.Line.ForeColor.RGB = RGB(84, 130, 53) Then
                    Shp.Line.ForeColor.RGB = 5921370 'else grey
                    color = "grey"
                Else
                    Shp.Line.ForeColor.RGB = 232132 'else whatever
                    color = "whatever"
                End If
            End If
        End If
    Next
    
    For Each Shp In ThisWorkbook.Sheets("Home").ChartObjects("HeatMap").Chart.Shapes
        If Left(Shp.Name, 5) = "Arrow" Then
            If Shp.Line.weight = wght Then
                If color = "blue" Then
                    Shp.Visible = True
                    Shp.Line.ForeColor.RGB = 16763030
                    Shp.Line.Transparency = Shp.Line.Transparency - 0.1 '' lowering the transparency allows to see better the highlighted arrows
                'Test new red
                ElseIf color = "red" Then
                    Shp.Line.ForeColor.RGB = RGB(255, 0, 0)
                    'shp.Line.Transparency = shp.Line.Transparency + 0.1
                    Shp.ZOrder msoBringToFront
                'Test new red
                ElseIf color = "green" Then
                    Shp.Line.ForeColor.RGB = RGB(84, 130, 53)
                    'shp.Line.Transparency = shp.Line.Transparency + 0.1
                    Shp.ZOrder msoBringToFront
                ElseIf color = "grey" Then
                    Shp.Line.ForeColor.RGB = 5921370
                    Shp.Line.Transparency = Shp.Line.Transparency + 0.1
                    Shp.ZOrder msoBringToFront
                ElseIf color = "whatever" Then
                    Shp.Visible = False
                End If
            End If
        End If
    Next Shp

End Sub


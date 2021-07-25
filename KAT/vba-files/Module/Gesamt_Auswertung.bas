Attribute VB_Name = "Gesamt_Auswertung"
Option Explicit



Sub scoring()

    Dim shPiv As Worksheet
    Dim tbl As ListObject
    Dim piv As PivotTable
    
    Dim derCol As Integer, bezCol As Integer, i As Integer, j As Integer, k As Integer
    Dim usedPart As Integer, devPart As Integer, rw As Integer, gesamt As Integer
    Dim derName As String
    Dim dataPivDer() As Variant, dataPivBez() As Variant, arr() As Variant
    
    Set shPiv = ThisWorkbook.Sheets("PIVOT")
    Set piv = shPiv.PivotTables("PivotTableMEGALISTE")
    Set tbl = ThisWorkbook.Sheets("Home").ListObjects("tab")
    
    piv.PivotFields("Kommunalität").ShowDetail = False
    
    '' clean table which will contain the scoring variables
    tbl.DataBodyRange.ClearContents
    
    With piv.PivotFields("Derivat")
        '' find dimension of arr (an array) --> how many selected derivat
        rw = selectedDerivatCount
        '' the 7 columns are : derivat name, x score, y score, n + nSA Anzahl, total Anzahl, used part Anzahl, developped parts Anzahl
        ReDim arr(1 To rw, 1 To 7)
        
        '' fill the array first column with the names of the valid derivat
        rw = 0
        For i = 1 To .PivotItems.count
            If .PivotItems(i).Visible = True Then
                rw = rw + 1
                derName = .PivotItems(i).Name
                arr(rw, 1) = derName
            End If
        Next i
        
        '' fill the 5th column with the total number of parts per derivat (depends on the filter chosen)
        dataPivDer = shPiv.Range(piv.TableRange1.Address)
        For i = 1 To rw
            derName = arr(i, 1)
            
            '' find column in which to look
            derCol = 0
            For j = 1 To UBound(dataPivDer, 2)
                If dataPivDer(2, j) = derName Then derCol = j: Exit For
            Next j
            
            If derCol = 0 Then
                gesamt = 0
            Else
                gesamt = dataPivDer(UBound(dataPivDer), derCol)
            End If
            
            arr(i, 5) = gesamt '' total of parts in a derivat
        Next i
    
        '' change the pivot and store the table in an array
        Call nnSAssSAKomAnzahlDer
        dataPivDer = shPiv.Range(piv.TableRange1.Address) '' we overwrite the first dataPivDer with another data
        
        '' change the pivot again and store it in another array
        Call ggSAKomAnzahlBez
        dataPivBez = shPiv.Range(piv.TableRange1.Address)
        
        For i = 1 To rw
            usedPart = 0: devPart = 0: derCol = 0: bezCol = 0
            derName = arr(i, 1)
            '' find the derivat column in the first array create above
            For j = 1 To UBound(dataPivDer, 2)
                If dataPivDer(2, j) = derName Then derCol = j: Exit For
            Next j
            
            If derCol = 0 Then
                devPart = 0
                usedPart = 0
            Else
                devPart = dataPivDer(UBound(dataPivDer), derCol)
                '' find theand Bezugsteil column in the second array create above
                For j = 1 To UBound(dataPivBez, 2)
                    If dataPivBez(2, j) = arr(i, 1) Then bezCol = j: Exit For
                Next j
                
                If bezCol = 0 Or InStr(derName, "(KP)") > 0 Then '' because Konfigprämissen derivat don't have Objekt-Name anyway.
                    usedPart = 0
                Else
                    '' compare unique parts between the two lists
                    '' one list is the list of developped parts for one derivat
                    '' the second list is the parts used in other derivat (as Bezugsteil)
                    For j = 3 To UBound(dataPivDer) - 1
                        '' the pivot display all Objekt-Name for all derivat
                        '' so we count as "developped parts" just the ones marked for a derivat, when the cell is not empty
                        If dataPivDer(j, derCol) <> vbNullString Then
                                '' go through the list of Objekt-Name in the Pivot
                                '' skip the "blue areas" like "Gesamtdergebnis" and "Zeilenbeschriftungen"
                                For k = 3 To UBound(dataPivBez) - 1
                                     If dataPivBez(k, bezCol) <> vbNullString Then
                                        '' if the Objekt-Name label is the same in the Bezugsteil list and the Derivatsteil list then it counts as one "used part"
                                        If dataPivDer(j, 1) = dataPivBez(k, 1) Then
                                            usedPart = usedPart + 1
                                            Exit For
                                        End If
                                     End If
                                Next k
                        End If
                    Next j
                End If
            End If
            
            '' store all results in the array
            arr(i, 6) = usedPart '' Anzahl from unique parts used in other derivats
            arr(i, 7) = devPart '' Anzahl n + nSA parts in a derivat
            arr(i, 4) = arr(i, 5) - arr(i, 7) '' Anzahl g + gSA + s + sSA
            
            '' the 3rd column is the actual y-axis score
            If devPart = 0 Then  '' here we want to avoid the division by 0
                arr(i, 3) = 0
            Else
                arr(i, 3) = arr(i, 6) / arr(i, 7)
            End If
            
            '' the 4th column is the actual x-axis score
            If arr(i, 5) = 0 Then '' here we want to avoid the division by 0
                arr(i, 2) = 0
            Else
                arr(i, 2) = arr(i, 4) / arr(i, 5)
            End If
        Next i
        
        '' write the array in the table
        tbl.Resize Range("tab[#All]").Resize(rw + 1, 7)
        tbl.DataBodyRange.Value = arr
    End With
    
    '' restore the pivot in it defaut setting "Gesamtdarstellung"
    Call GesamtPivot

    piv.PivotFields("Kommunalität").ShowDetail = True
    
End Sub



Sub createScoringMap()
    
    Dim ws As Worksheet
    Dim objChrt As ChartObject
    Dim s As Series
    Dim lngIndex As Integer, Derivat As String
    Dim Shp As Shape

    Set ws = ThisWorkbook.Sheets("Home")
    With ws
        On Error Resume Next
        .ChartObjects("ScoringDia").Delete
        On Error GoTo 0

        Set objChrt = .ChartObjects.add(180, 1200, 600, 450)
        With objChrt.Chart
            .Parent.Name = "ScoringDia"
            Set s = .SeriesCollection.NewSeries()
            With s
                .ChartType = xlXYScatter
                .Name = "Scoring"
                .XValues = ws.ListObjects("tab").DataBodyRange.Columns(2)
                .Values = ws.ListObjects("tab").DataBodyRange.Columns(3)
                For lngIndex = 1 To .Points.count
                    Derivat = ws.ListObjects("tab").DataBodyRange(lngIndex, 1)
                    .Points(lngIndex).HasDataLabel = True
                    .Points(lngIndex).DataLabel.Text = Derivat
                Next lngIndex
            End With
        End With
    End With
    
    Call formatScoring
    
End Sub

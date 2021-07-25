Attribute VB_Name = "Pivot_Detail"
Option Explicit
Sub initializeDetail()
    
    '' this sub is quite long (for big Gesamtdarstllung, it can last several seconds)
    '' the goal here is to assign a macro to each arrow --> on action and pass arguments (the arrow name which states its start and finish )
    Dim Shp As Shape
    If ThisWorkbook.Sheets("Home").detailMode <> True Then Exit Sub
    
    For Each Shp In ThisWorkbook.Sheets("Home").ChartObjects("HeatMap").Chart.Shapes
        If Left(Shp.Name, 5) = "Arrow" Then
            If Shp.OnAction <> vbNullString Then Exit Sub
            Shp.OnAction = "'detailGesamt """ & Shp.Name & """'"
        End If
    Next Shp
End Sub


Sub detailEinzel(ByVal Target As Range)

    Dim der As String, slstr As String
    Dim Kom As String
    Dim i As Integer, j As Integer
    Dim piv As PivotTable
    Dim zusTbl As ListObject
    
    slstr = SelectedSlicer
    
    '' if the place on which the user has double clicked is in the einzeldarstellung table, then build the table with the details by going in the pivot
    '' in the pivot, triger the "show details" on the appropriate cell
    Set piv = ThisWorkbook.Sheets("PIVOT").PivotTables("PivotTableMEGALISTE")
    Set zusTbl = ThisWorkbook.Sheets("Home").ListObjects("ZusTab")
    
    freezeTriger = True
    piv.ManualUpdate = True
    
    For i = 1 To piv.PivotFields.count
        If piv.PivotFields(i).Name = "HZ1" Then 'column
            piv.PivotFields(i).Orientation = xlHidden
        ElseIf piv.PivotFields(i).Name = "HZ2" Then 'column
            piv.PivotFields(i).Orientation = xlHidden
        ElseIf piv.PivotFields(i).Name = "HZ3" Then 'column
            piv.PivotFields(i).Orientation = xlHidden
        End If
    Next i
    
    For i = 1 To 6
        If Target.Address = zusTbl.DataBodyRange(1, i).Address Then
            If slstr <> Right(ThisWorkbook.Sheets("Home").Range("AO13").Value, Len(slstr)) Then
                MsgBox "The current filter selection and the Einzeldarstellung filter selection do not match." & vbNewLine & "Please reselect the same filters to display details."
                Exit Sub
            End If
            
            If piv.PivotFields("Kommunalität").Orientation <> xlColumnField Or _
                piv.PivotFields("Fzg.typ Bezugsteil").Orientation <> xlRowField Or _
                piv.DataFields(1).Name <> "Anzahl von Kommunalität" Then
                MsgBox "The pivot was not in the Einzeldarstellung mode." & vbNewLine & "Please regenerate the Einzeldarstellung view."
                Exit Sub
            End If
            
            Kom = zusTbl.HeaderRowRange(i).Value '' this identifies the Kommunalität (g, gSA, etc )
            der = ThisWorkbook.Sheets("Home").ChartObjects("pieDia").Chart.ChartTitle.Caption
            Call deleteWs("Detail_" & der & "_" & Kom)
            '' this is the appropriate cell
            
            piv.ManualUpdate = False
            ThisWorkbook.Sheets("PIVOT").Visible = True
            ThisWorkbook.Sheets("PIVOT").Activate
            
            With piv.PivotFields("Fzg.typ Bezugsteil")
                .EnableMultiplePageItems = False
                .ClearAllFilters
                .EnableMultiplePageItems = True
            End With
        
            piv.PivotFields("Kommunalität").PivotItems(Kom).DataRange.Cells(piv.DataBodyRange.Rows.count).ShowDetail = True
            ThisWorkbook.Sheets("PIVOT").Visible = False
            
            With ActiveSheet
                .Name = "Detail_" & der & "_" & Kom
                .Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                .Range("A1").Value = .Name & " | " & ThisWorkbook.Sheets("Home").Range("AO13").Value
                .Move after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
            End With
            
            For j = 1 To piv.PivotFields.count
                If piv.PivotFields(j).Name = "HZ1" Then 'column
                    piv.PivotFields(j).Orientation = xlColumnField
                    piv.PivotFields(j).Position = 2
                ElseIf piv.PivotFields(j).Name = "HZ2" Then 'column
                    piv.PivotFields(j).Orientation = xlColumnField
                    piv.PivotFields(j).Position = 3
                ElseIf piv.PivotFields(j).Name = "HZ3" Then 'column
                    piv.PivotFields(j).Orientation = xlColumnField
                    piv.PivotFields(j).Position = 4
                End If
            Next j
            
            freezeTriger = False
            piv.ManualUpdate = False
            
            Exit Sub
        End If
    Next i
End Sub



Sub detailGesamt(str As String)

    Dim arr() As String, slstr As String
    Dim fzg As String, der As String
    Dim sh As Worksheet
    Dim piv As PivotTable
    
    '' get the start and finish of the arrow from its name which was passed as argument
    '' from there
    slstr = SelectedSlicer
    If ThisWorkbook.Sheets("Home").detailMode = True Then
        If slstr <> Right(ThisWorkbook.Sheets("Home").Range("A41").Value, Len(slstr)) Then
            MsgBox "The current filter selection and the Gesamtdarstellung filter selection do not match." & vbNewLine & "Please reselect the same filters to display details."
            Exit Sub
        End If
        
        Set sh = ThisWorkbook.Sheets("PIVOT")
        Set piv = sh.PivotTables("PivotTableMEGALISTE")
        
        If piv.PivotFields("Derivat").Orientation <> xlColumnField Or _
            piv.PivotFields("Fzg.typ Bezugsteil").Orientation <> xlRowField Or _
            piv.DataFields(1).Name <> "Anzahl von Kommunalität" Then
            MsgBox "The pivot was not in the Gesamtdarstellung mode." & vbNewLine & "Please regenerate the Gesamtdarstellung view."
            Exit Sub
        End If
        
        str = Replace(str, "Arrow", vbNullString)
        arr() = Split(str, "-")
        fzg = arr(0) '' source
        der = arr(1) '' target
        Call deleteWs("Detail_" & str)
        '' this is the appropriate cell
        Intersect(piv.PivotFields("Fzg.typ Bezugsteil").PivotItems(fzg).DataRange.EntireRow, piv.PivotFields("Derivat").PivotItems(der).DataRange).ShowDetail = True
        
        '' change a few things on the newly created worksheet
        '' display name and filters to identify it
        '' place it a the end of the workbook
        With ActiveSheet
            .Name = "Detail_" & str
            .Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            .Range("A1").Value = .Name & " | " & ThisWorkbook.Sheets("Home").Range("A41").Value
            .Move after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        End With
    End If
End Sub

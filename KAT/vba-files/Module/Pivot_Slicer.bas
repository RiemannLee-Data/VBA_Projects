Attribute VB_Name = "Pivot_Slicer"
Option Explicit

Sub updSl(Name As String)
    '' update slicers
    '' creates if they do not exit
    '' deletes if they do
    ''
    freezeTriger = True '' disable pivot update event
    
    Dim shPiv As Worksheet
    Dim piv As PivotTable
    Dim i As Integer, anzahl As Integer
    Dim found As Boolean

    With ThisWorkbook
        Set shPiv = .Sheets("PIVOT")
        Set piv = shPiv.PivotTables("PivotTableMEGALISTE")
        anzahl = .SlicerCaches.count
        
        '' loop through the slicers, if the slicer already exists, then delete it
        If anzahl > 0 Then
            For i = anzahl To 1 Step -1
                If .SlicerCaches(i).Slicers(1).Caption = Name Then
                    .SlicerCaches(i).Delete
                    piv.PivotFields(Name).ClearAllFilters
                    found = True
                    Exit For
                End If
            Next i
        End If
        
        '' if it doesn't exist then create it
        If found = False Then
            .SlicerCaches.Add2(piv, Name).Slicers.add .Sheets("Home"), , Name, Name, 10, 180 + anzahl * 140, 135, 165
        End If
    End With
    
    freezeTriger = False
    
End Sub


Sub clearSlicer()

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    freezeTriger = True
    
    '' clear all existing slicers and their selected fields
    Dim contr As Control
    Dim slCaches As SlicerCaches
    Dim slCache As SlicerCache

    Set slCaches = ThisWorkbook.SlicerCaches

    For Each slCache In slCaches
            slCache.Delete
    Next slCache
    
    Call resetSlicer
    
    freezeTriger = False
    
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With

End Sub

Sub resetSlicer()

    freezeTriger = True
    
    '' clear all existing slicers and their selected fields
    Dim piv As PivotTable
    Dim pf As pivotfield
    With ThisWorkbook.Sheets("PIVOT")
        If .PivotTables.count > 0 Then
            Set piv = .PivotTables("PivotTableMEGALISTE")
            For Each pf In piv.PivotFields
                    pf.ClearAllFilters
            Next pf
        End If
    End With
    
    freezeTriger = False

End Sub

Function SelectedSlicer() As String

    Dim slName As String
    Dim str As String
    Dim i As Integer, j As Long
    Dim count As Integer, max As Integer
    
    '' loop through slicer and slicer items and store names in a string
    With ThisWorkbook
        For i = 1 To .SlicerCaches.count
            If .SlicerCaches(i).SlicerItems.count > .SlicerCaches(i).VisibleSlicerItems.count Then
                slName = Replace(Replace(.SlicerCaches(i).Name, "Datenschnitt_", vbNullString), "_", " ")
                str = str & slName & ": "
                For j = 1 To .SlicerCaches(i).VisibleSlicerItems.count
                    str = str & .SlicerCaches(i).VisibleSlicerItems(j).Name & ", "
                Next j
                str = Left(str, Len(str) - 2) & " | "
            End If
        Next i
    End With
    
    If str = vbNullString Then
        str = "No Filters|"
    End If
    
    str = cleanValue(str)
    str = Left(str, Len(str) - 1)
    str = "Filtereinstellung | " & str
    
    SelectedSlicer = str
    
End Function

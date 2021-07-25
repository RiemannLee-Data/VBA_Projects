Attribute VB_Name = "Pivot_Switch"
Option Explicit
Public freezeTriger As Boolean



Sub createPivot()

    '' this sub recreates the pivot when a new derivat is added or when one is deleted
    Dim wb As Workbook
    Dim shPiv As Worksheet, shMega As Worksheet, shPivFB As Worksheet
    Dim sourcedata As String
    Dim lastRowMega As Long, lastColMega As Integer
    Dim Ret As Boolean
    
    '' the slicers are deleted because they are linked to the current pivot table
    Call clearSlicer
    
    '' we have a "on pivot change" event on the Pivot worksheet
    '' this allows us to turn it off, just the time to change the pivot
    freezeTriger = True
    
    Set shPiv = ThisWorkbook.Sheets("PIVOT")
    Set shPivFB = ThisWorkbook.Sheets("PIVOT_FB")

    '' delete current pivot from shpiv
    shPiv.Cells.Clear
    shPivFB.Cells.Clear
    
    '' check if megaliste is open, if not, then open it
    Ret = IsWorkBookOpen(ThisWorkbook.Path & "\KAT_Vorlage\MEGALISTE.xlsx")
    If Ret <> True Then Workbooks.Open (ThisWorkbook.Path & "\KAT_Vorlage\MEGALISTE.xlsx")
    
    '' get last row and last column of the megaliste
    Set wb = Workbooks("MEGALISTE.xlsx")
    Set shMega = wb.Sheets("Derivat")
    lastRowMega = shMega.UsedRange.Rows.count
    lastColMega = shMega.UsedRange.Columns.count
    
    '' add last row and last column in the string that is used to show the pivot source range
    If lastRowMega > 1 Then
        sourcedata = ThisWorkbook.Path & "\KAT_Vorlage\[MEGALISTE.xlsx]Derivat!R1C1:R" & lastRowMega & "C" & lastColMega
        
        ThisWorkbook.PivotCaches.create(SourceType:=xlDatabase, _
        sourcedata:=sourcedata, _
        Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:=shPiv.Name & "!R3C1", _
        TableName:="PivotTableMEGALISTE", DefaultVersion:=xlPivotTableVersion15
        '' put pivot in it's default configuration : Gesamtdarstellung
        Call GesamtPivot
        Call FBPiv
    End If
    
    freezeTriger = False
    
    shPiv.Visible = False
End Sub



Sub EinzelPivot()

    '' sets up the pivot as it can used for the Einzeldarstellung
    Dim piv As PivotTable
    Dim i As Integer, count  As Integer
    Set piv = ThisWorkbook.Sheets("PIVOT").PivotTables("PivotTableMEGALISTE")
    
    With piv
        freezeTriger = True
        .ManualUpdate = True
        
            count = .DataFields.count
            If count = 0 Then
                .AddDataField .PivotFields("Kommunalität"), "Anzahl von Kommunalität", xlCount
            Else
                For i = count To 1 Step -1
                    If .DataFields(i).Name <> "Anzahl von Kommunalität" Then
                        .DataFields(i).Orientation = xlHidden
                    End If
                Next i
            End If
            
            .PivotFields("Kommunalität").ClearAllFilters
            For i = 1 To .PivotFields.count
                If .PivotFields(i).Name = "Fzg.typ Bezugsteil" And .PivotFields(i).Orientation <> 1 Then 'row
                    .PivotFields(i).Orientation = xlRowField
                ElseIf .PivotFields(i).Name = "Kommunalität" And .PivotFields(i).Orientation <> 2 Then 'column
                    .PivotFields(i).Orientation = xlColumnField
                    .PivotFields(i).Position = 1
                ElseIf .PivotFields(i).Name = "HZ1" And .PivotFields(i).Orientation <> 2 Then 'column
                    .PivotFields(i).Orientation = xlColumnField
                    .PivotFields(i).Position = 2
                ElseIf .PivotFields(i).Name = "HZ2" And .PivotFields(i).Orientation <> 2 Then 'column
                    .PivotFields(i).Orientation = xlColumnField
                    .PivotFields(i).Position = 3
                ElseIf .PivotFields(i).Name = "HZ3" And .PivotFields(i).Orientation <> 2 Then 'column
                    .PivotFields(i).Orientation = xlColumnField
                    .PivotFields(i).Position = 4
                ElseIf .PivotFields(i).Orientation <> xlPageField And _
                    .PivotFields(i).Name <> "Kommunalität" And _
                    .PivotFields(i).Name <> "Fzg.typ Bezugsteil" And _
                    .PivotFields(i).Name <> "HZ1" And _
                    .PivotFields(i).Name <> "HZ2" And _
                    .PivotFields(i).Name <> "HZ3" Then
                    .PivotFields(i).Orientation = xlPageField
                End If
            Next i
        
        '' calculate the table content and sort by Kommunalität
        .PivotFields("Fzg.typ Bezugsteil").AutoSort xlDescending, "Anzahl von Kommunalität"
        
        .ManualUpdate = False
        freezeTriger = False
    End With
End Sub



Sub GesamtPivot()

    Dim piv As PivotTable
    Dim count As Integer, i As Integer
    Set piv = ThisWorkbook.Sheets("PIVOT").PivotTables("PivotTableMEGALISTE")
    '' sets up the pivot as it can used used for the Gesamtdarstellung
    With piv
        freezeTriger = True
        .ManualUpdate = True
        
            count = .DataFields.count
            If count = 0 Then
                .AddDataField .PivotFields("Kommunalität"), "Anzahl von Kommunalität", xlCount
            Else
                For i = count To 1 Step -1
                    If .DataFields(i).Name <> "Anzahl von Kommunalität" Then
                        .DataFields(i).Orientation = xlHidden
                    End If
                Next i
            End If
            
            .PivotFields("Kommunalität").ClearAllFilters
            For i = 1 To .PivotFields.count
                If .PivotFields(i).Name = "Derivat" And .PivotFields(i).Orientation <> 2 Then '' 2 = column field
                    .PivotFields(i).Orientation = xlColumnField
                    .PivotFields(i).AutoSort xlAscending, "Derivat"
                ElseIf .PivotFields(i).Name = "Fzg.typ Bezugsteil" And .PivotFields(i).Orientation <> 1 Then '' 1 = row field
                    .PivotFields(i).Orientation = xlRowField
                ElseIf .PivotFields(i).Orientation <> xlPageField And _
                    .PivotFields(i).Name <> "Derivat" And _
                    .PivotFields(i).Name <> "Fzg.typ Bezugsteil" Then
                    .PivotFields(i).Orientation = xlPageField '' filter field
                End If
            Next i
        
        .ManualUpdate = False
        freezeTriger = False
    End With

End Sub



Sub FBPiv()

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With

    Dim wbk As Workbook
    Dim piv As PivotTable
    
    For Each wbk In Workbooks
        If InStr(wbk.Name, "KAT") Then
        wbk.Activate
            Exit For
        End If
    Next wbk

    ThisWorkbook.Sheets("PIVOT_FB").Visible = True
    ThisWorkbook.Sheets("PIVOT").Visible = True
    ThisWorkbook.Sheets("PIVOT_FB").Cells.Clear
    ThisWorkbook.Sheets("PIVOT").Activate
    ThisWorkbook.Sheets("PIVOT").PivotTables("PivotTableMEGALISTE").PivotSelect "", xlDataAndLabel, True
    Selection.Copy
    Sheets("PIVOT_FB").Activate
    Range("A1").Select
    ActiveSheet.Paste
    For Each piv In ActiveSheet.PivotTables
        piv.Name = "PivotTableFB"
    Next
    
    Set piv = ActiveSheet.PivotTables("PivotTableFB")
    piv.PivotSelect "", xlDataAndLabel, True
    piv.PivotFields("Komponente").Orientation = xlHidden
    piv.PivotFields("Objekt-Name").Orientation = xlHidden
    piv.PivotFields("FB").Orientation = xlHidden
    piv.PivotFields("Modulorg.").Orientation = xlHidden
    piv.PivotFields("techn. Beschr.").Orientation = xlHidden
    piv.PivotFields("Kom. Erstverwendung").Orientation = xlHidden
    piv.PivotFields("Fzg.typ Erstverw.").Orientation = xlHidden
    piv.PivotFields("Fzg.typ Bezugsteil").Orientation = xlHidden
    piv.PivotFields("Beziehungswissen").Orientation = xlHidden
    piv.PivotFields("BK-Cluster Text Variante").Orientation = xlHidden
    piv.PivotFields("BK-Cluster Variante").Orientation = xlHidden
    piv.PivotFields("Group Baukastenbezeichnung").Orientation = xlHidden
    piv.PivotFields("Produkt Baukasten").Orientation = xlHidden
    piv.PivotFields("Archetyp-Kateg. Variante").Orientation = xlHidden
    piv.PivotFields("Prozessbaukasten").Orientation = xlHidden
    piv.PivotFields("BB Teilekomm.").Orientation = xlHidden
    piv.PivotFields("Anzahl von Kommunalität").Orientation = xlHidden
    piv.PivotFields("Architekturkennzeichen").Orientation = xlHidden
    piv.PivotFields("PosVar-GUID").Orientation = xlHidden
    piv.PivotFields("Dimensionslosekommunalitaet").Orientation = xlHidden
    piv.PivotFields("HZ1").Orientation = xlHidden
    piv.PivotFields("Knoten").Orientation = xlHidden
    piv.PivotFields("HZ2").Orientation = xlHidden
    piv.PivotFields("HZ3").Orientation = xlHidden
    piv.PivotFields("PPG").Orientation = xlHidden
    
    piv.AddDataField piv.PivotFields("Kommunalität"), "Anzahl von Kommunalität", xlCount
    
    With piv.PivotFields("Derivat")
        .Orientation = xlPageField
        .ClearAllFilters
    End With
    
    With piv.PivotFields("Kommunalität")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    With piv.PivotFields("FB")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ThisWorkbook.Sheets("PIVOT_FB").Visible = True
    ThisWorkbook.Sheets("PIVOT").Visible = True

End Sub


Sub nnSAssSAKomAnzahlDer() '' this is called during the Auswertung filter NT and ST

    Dim piv As PivotTable
    Dim i As Integer, j As Integer
    Set piv = ThisWorkbook.Sheets("PIVOT").PivotTables("PivotTableMEGALISTE")
    
    With piv
        freezeTriger = True
        .ManualUpdate = True
        
        With .PivotFields("Fzg.typ Bezugsteil")
            .Orientation = xlPageField
        End With
    
        With .PivotFields("Kommunalität")
            For i = 1 To .PivotItems.count
                If .PivotItems(i).Name = "n" Or .PivotItems(i).Name = "nSA" Or .PivotItems(i).Name = "s" Or .PivotItems(i).Name = "sSA" Then
                    .PivotItems(i).Visible = True
                Else
                    .PivotItems(i).Visible = False
                End If
            Next i
        End With
        
        With .PivotFields("Objekt-Name")
            .Orientation = xlRowField
        End With
        
        .ManualUpdate = False
        freezeTriger = False
        
    End With
End Sub



Sub ggSAKomAnzahlBez() '' this is called during the Auswertung filter GT

    Dim piv As PivotTable
    Dim i As Integer
    Set piv = ThisWorkbook.Sheets("PIVOT").PivotTables("PivotTableMEGALISTE")
    
    With piv
        .ManualUpdate = True
        freezeTriger = True
    
        With .PivotFields("Derivat")
            .Orientation = xlPageField
        End With
        
        With .PivotFields("Fzg.typ Bezugsteil")
            .Orientation = xlColumnField
        End With
        
        With .PivotFields("Kommunalität")
            .ClearAllFilters
            For i = 1 To .PivotItems.count
                If .PivotItems(i).Name = "g" Or .PivotItems(i).Name = "gSA" Then
                    .PivotItems(i).Visible = True
                Else
                    .PivotItems(i).Visible = False
                End If
            Next i
        End With
        
        freezeTriger = False
        .ManualUpdate = False
    End With
End Sub

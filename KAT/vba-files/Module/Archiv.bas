Attribute VB_Name = "Archiv"
Sub archivieren(str As String)

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With

    Dim Derivat As String
    Dim z As Variant

    aktuellesQuartal = DatePart("q", Date)
    aktuellesJahr = Right(Date, 4)
    aktuellesQuartal = aktuellesQuartal & ". Quartal " & aktuellesJahr

    If derCount = 1 Then
        Call archivierenEinzel(str)
    ElseIf derCount > 1 Then
        z = Split(str, ",")
        For i = 0 To UBound(z)
            Derivat = z(i)
            Call EinzelPivot
            Call archivierenEinzel(Derivat)
            
            Set Startrange = ThisWorkbook.Worksheets("Typschl").UsedRange.Find(Derivat, LookIn:=xlValues, Lookat:=xlWhole)
            row = Startrange.row
            col = Startrange.Column
            For j = row To ThisWorkbook.Worksheets("Typschl").UsedRange.Rows.count
                If Not IsEmpty(ThisWorkbook.Worksheets("Typschl").Cells(j, 2)) And Not IsEmpty(ThisWorkbook.Worksheets("Typschl").Cells(j, 7)) Then
                    ThisWorkbook.Worksheets("Typschl").Cells(j, 12) = aktuellesQuartal
                    Exit For
                End If
            Next j
        Next i
    Else
        MsgBox ("Bitte drücken Sie den Refresh-Button.")
        Exit Sub
    End If
    
    If MsgBox("Möchten Sie die Historie anzeigen?", vbYesNo) = vbNo Then
        Workbooks("HISTORIE.xlsx").Close True
    End If
    
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With

    
End Sub



Sub archivierenEinzel(str As String)
    
    Dim pfadM As String, pfadH As String, Kom As String
    Dim wbMegaM As Workbook, wbMegaH As Workbook, wbKAT As Workbook
    Dim shMegaM As Worksheet, shMegaH As Worksheet, WSsucheH As Worksheet, ws1 As Worksheet, wsNeu As Worksheet, wsHome As Worksheet, wsPiv As Worksheet
    Dim vorhanden As Boolean, neuesQuartal As Boolean
    Dim dataMega() As Variant, anz_Teile() As Variant
    Dim Startrange As Range
    Dim col As Integer, j As Integer, maxi As Long, E As Long, i As Long
        
    Set wbKAT = ActiveWorkbook
    Set wsHome = wbKAT.Worksheets("Home")
    Set wsPiv = wbKAT.Worksheets("PIVOT")
        
    'öffnen Historie
    pfadH = ThisWorkbook.Path & "\KAT_Vorlage\HISTORIE.xlsx"
    If IsWorkBookOpen(pfadH) <> True Then
        Workbooks.Open (pfadH)
    Else
        Workbooks("HISTORIE.xlsx").Activate
    End If
    
    Set wbMegaH = Workbooks("HISTORIE.xlsx")
        
    vorhanden = False
    
    For Each WSsucheH In wbMegaH.Worksheets
        If UCase(WSsucheH.Name) = UCase(str) Then
            vorhanden = True
            Set wsNeu = WSsucheH
            Exit For
        End If
    Next WSsucheH
    
    If vorhanden = False Then
        wbMegaH.Activate
        wbMegaH.Worksheets("Vorlage").Copy before:=Sheets(1)
        wbMegaH.Worksheets("Vorlage (2)").Name = str
        Set wsNeu = wbMegaH.Worksheets(str)
        wsNeu.Visible = True
        wsNeu.Activate
    End If
        
    ReDim anz_Teile(1 To 6)
    i = 1
    wsPiv.Activate
    'PIVOT auf aktuelles Derivat anwenden
    With wsPiv.PivotTables("PivotTableMEGALISTE").PivotFields("Derivat")
        .EnableMultiplePageItems = False
        .ClearAllFilters
        .EnableMultiplePageItems = True
        For Each pivItem In .PivotItems
            If pivItem = str Then
                pivItem.Visible = True
            Else
                pivItem.Visible = False
            End If
        Next
   End With
   
   '#######################################################
   'Was wenn keine SA Teile vorhanden sind
   
    'Daten aus Pivot direkt nehmen!
    Set Startrange = Nothing
    Set Startrange = wsPiv.Columns(1).Find("Gesamtergebnis", LookIn:=xlValues, Lookat:=xlWhole)
    ergRow = Startrange.row
    Set Startrange = Nothing
        Set Startrange = wsPiv.UsedRange.Find("g Ergebnis", LookIn:=xlValues, Lookat:=xlWhole)
        If Startrange Is Nothing Then
            col_g = 0
        Else
            col_g = Startrange.Column
        End If
    Set Startrange = Nothing
        Set Startrange = wsPiv.UsedRange.Find("gSA Ergebnis", LookIn:=xlValues, Lookat:=xlWhole)
        If Startrange Is Nothing Then
            col_gSA = 0
        Else
            col_gSA = Startrange.Column
        End If
    Set Startrange = Nothing
        Set Startrange = wsPiv.UsedRange.Find("n Ergebnis", LookIn:=xlValues, Lookat:=xlWhole)
        If Startrange Is Nothing Then
            col_n = 0
        Else
            col_n = Startrange.Column
        End If
    Set Startrange = Nothing
        Set Startrange = wsPiv.UsedRange.Find("nSA Ergebnis", LookIn:=xlValues, Lookat:=xlWhole)
        If Startrange Is Nothing Then
            col_nSA = 0
        Else
            col_nSA = Startrange.Column
        End If
    Set Startrange = Nothing
        Set Startrange = wsPiv.UsedRange.Find("s Ergebnis", LookIn:=xlValues, Lookat:=xlWhole)
        If Startrange Is Nothing Then
            col_s = 0
        Else
            col_s = Startrange.Column
        End If
    Set Startrange = Nothing
        Set Startrange = wsPiv.UsedRange.Find("sSA Ergebnis", LookIn:=xlValues, Lookat:=xlWhole)
        
        If Startrange Is Nothing Then
            col_sSA = 0
        Else
            col_sSA = Startrange.Column
        End If

        If col_gSA = 0 Then
            anz_Teile(1) = wsPiv.Cells(ergRow, col_g)
        Else
            anz_Teile(1) = wsPiv.Cells(ergRow, col_g) + wsPiv.Cells(ergRow, col_gSA)
        End If
        
        If col_nSA = 0 Then
            anz_Teile(3) = wsPiv.Cells(ergRow, col_n)
        Else
            anz_Teile(3) = wsPiv.Cells(ergRow, col_n) + wsPiv.Cells(ergRow, col_nSA)
        End If
        
        If col_sSA = 0 Then
            anz_Teile(2) = wsPiv.Cells(ergRow, col_s)
        Else
            anz_Teile(2) = wsPiv.Cells(ergRow, col_s) + wsPiv.Cells(ergRow, col_sSA)
        End If

    wsNeu.Activate
    Set Startrange = wsNeu.UsedRange.Find("1. Quartal 2017", LookIn:=xlValues, Lookat:=xlWhole)
    row = Startrange.row
    
    aktuellesQuartal = DatePart("q", Date)
    aktuellesJahr = Right(Date, 4)
    aktuellesQuartal = aktuellesQuartal & ". Quartal " & aktuellesJahr
    
    Set Startrange = Nothing
    Set Startrange = wsNeu.UsedRange.Find(aktuellesQuartal, LookIn:=xlValues, Lookat:=xlWhole)

    neuesQuartal = False
    If Startrange Is Nothing Then
    'Neues Quartal am Ende hinzufügen
        neuesQuartal = True
        col = wsNeu.UsedRange.Columns.count + 1
    Else
        col = Startrange.Column
    End If
    
    For i = 1 To 3
    wsNeu.Cells(row + i, col) = anz_Teile(i)
    Next i
    
    If neuesQuartal = True Then
        wsNeu.Cells(row, col) = aktuellesQuartal
        wsNeu.Cells(row + 1, col).Interior.color = RGB(112, 173, 71)
        wsNeu.Cells(row + 2, col).Interior.color = RGB(255, 255, 0)
        wsNeu.Cells(row + 3, col).Interior.color = RGB(255, 0, 0)
        ActiveSheet.ChartObjects("Diagramm 1").Activate
        ActiveChart.PlotArea.Select
        ActiveChart.SetSourceData source:=wsNeu.Range(Cells(row, 2), Cells(row + 3, col))
    End If

    wbKAT.Activate
    
End Sub

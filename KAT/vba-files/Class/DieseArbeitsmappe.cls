VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
'Sub zum archivieren der Derivate, wird quartalsweise abgefragt

    '' Deklaration der Varaible
    Dim wb As Workbook
    Dim ws1 As Worksheet
    
    '' wenn man mit einem bestimmten Objekt mehrere Dinge auf einmal machen möchte
    '' um den Code zu vereinfachen, gibt es dafür With-Anweisung
    '' alles hier sind Eigenschaften
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With

    '' Obejektzuweisung erfordern die Set-Anweisung
    Set wb = ActiveWorkbook
    Set ws1 = wb.Worksheets("Home")

    aktuellesQuartal = DatePart("q", Date)
    aktuellesJahr = Right(Date, 4)
    aktuellesQuartal = aktuellesQuartal & ". Quartal " & aktuellesJahr
    
    With ThisWorkbook.Sheets("PIVOT").PivotTables("PivotTableMEGALISTE")
        AnzahlDerivategesamt = .PivotFields("Derivat").PivotItems.count
    End With
    
    Set Startrange = ws1.UsedRange.Find(aktuellesQuartal, LookIn:=xlValues, Lookat:=xlWhole)
    
    If Startrange Is Nothing Then
        Exit Sub
    End If
    
    row = Startrange.row
    col = Startrange.Column
    
    wb.Worksheets("Typschl").Visible = True
    wb.Worksheets("Typschl").Activate
    
    anzahl = Application.WorksheetFunction.CountIf(wb.Worksheets("Typschl").Range(Cells(1, 12), Cells(wb.Worksheets("Typschl").UsedRange.Rows.count, 12)), aktuellesQuartal)
    
    If anzahl = AnzahlDerivategesamt Or anzahl > AnzahlDerivategesamt Then
        ws1.Cells(row, col + 1) = "archiviert"
    ElseIf anzahl < AnzahlDerivategesamt Then
        ws1.Cells(row, col + 1) = ""
        MsgBox ("Alle Derivate sollten quartalsweise archiviert werden. Für das aktuelle Quartal ( " & aktuellesQuartal & " ) liegen noch keine Daten vor." & vbCrLf & "Bitte löschen Sie alle Derivate und importieren die aktuellen Strukturberichte.")
    End If
    
    wb.Worksheets("Typschl").Visible = False
    
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With


End Sub

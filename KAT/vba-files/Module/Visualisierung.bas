Attribute VB_Name = "Visualisierung"
Public strFile1 As String
Public FilePath As String
Public Gültigkeitsdatum As String


Sub Datenimport()

    Dim wk1 As Workbook
    Dim sh1 As Worksheet, sh2 As Worksheet, ws1 As Worksheet, ws2 As Worksheet
    Dim Treffer() As String, Derivat As String, Daten As String, Gültigkeitsdatum As String, Kürzen1 As String, Kürzen2 As String
    Dim Position As Integer, Position2 As Integer, Spaltenzahl As Integer, Richtig As Integer, Falsch As Integer, i As Integer, j As Integer, k As Integer, m As Integer, g As Integer, s As Integer, n As Integer
    Dim Anpassung As Boolean
    Dim Zeilenanzahl As Integer 'Dim Zeilenanzahl As Long
    
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    Vis_weiter = False
    
    'Prüfen, ob genau 1 Derivat ausgewählt ist
    'Bei Fehlermeldung obwohl Derivat angewählt ist, den Filter löschen und neu hinzufügen oder "Refresh"-Button klicken
    If der = "" Then
        MsgBox ("Bitte wählen Sie ein Derivat an.")
        Exit Sub
    ElseIf derCount > 1 Then
        MsgBox ("Bitte wählen Sie nur ein Derivat aus.")
        Exit Sub
    End If
    
    Derivat = der
    
    'Hauptordner erstellen
    FilePath = ThisWorkbook.Path & "\KAT_Vorlage\05_Visualisierung"
    strFile1 = FilePath & "\Visualisierung_" & Derivat
    If Dir(strFile1, vbDirectory) <> "" Then
    File = MsgBox("Der Ordner " & strFile1 & vbCrLf & "ist bereits angelegt. Soll dieser gelöscht werden?", vbYesNoCancel + vbCritical)
    
    If File = 7 Then
        MsgBox "Der Vorgang wird abgebrochen.", vbExclamation
        Call Shell("Explorer /e, " & FilePath, vbNormalFocus)
        Exit Sub
    ElseIf File = 6 Then
        'Ordner in Visualisierung löschen
        Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        objFSO.DeleteFolder (strFile1)
        
        Set objFSO = Nothing
        Application.Wait Now + TimeSerial(0, 0, 2)
        MkDir strFile1
    ElseIf File = 2 Then
        Exit Sub
        End If
    Else
        MkDir strFile1
    End If

    
    'Worksheets einblenden
    For Each Sheet In ActiveWorkbook.Worksheets
    Sheet.Visible = True
    Next Sheet

    Set ws1 = ThisWorkbook.Worksheets("VIS_MAKRO")
    Set ws2 = ThisWorkbook.Worksheets("VIS_PIVOT")
    Set wk1 = Workbooks.Open(ThisWorkbook.Path & "\KAT_Vorlage\MEGALISTE.xlsx")
    Set sh1 = wk1.Worksheets("Derivat")
    
    If sh1.FilterMode Then sh1.ShowAllData

    'Prüfen, ob noch Daten in den Sheets stehen
    If Not IsEmpty(ws1.Cells(2, 1)) Then
        Call Löschen
    ElseIf Not IsEmpty(ws2.Cells(2, 1)) Then
        Call Löschen
    End If

    sh1.Activate
    Zeilenanzahl = sh1.Cells(Rows.count, 1).End(xlUp).row
    Spaltenzahl = sh1.Cells(1, Columns.count).End(xlToLeft).Column
    sh1.Range(Cells(1, 1), Cells(Zeilenanzahl, Spaltenzahl)).AutoFilter Field:=1, Criteria1:=der
    
    'Kopieren aller relevanten Daten
    Zeilenanzahl = sh1.Cells(Rows.count, 1).End(xlUp).row
    ReDim Treffer(1 To Spaltenzahl)

    'RELEVANTE DATEN FÜR VISUALISIERUNG
    'In Ausleitung Spalten finden
     For i = 1 To Spaltenzahl
         Treffer(i) = sh1.Cells(1, i)
         If Treffer(i) = "Modulorg." Then
             sh1.Activate
             sh1.Range(Cells(2, i), Cells(Zeilenanzahl, i)).SpecialCells(xlCellTypeVisible).Copy
             ws1.Activate
             ws1.Cells(Rows.count, "B").End(xlUp).offset(1, 0).PasteSpecial
         ElseIf Treffer(i) = "PPG" Then
             sh1.Activate
             sh1.Range(Cells(2, i), Cells(Zeilenanzahl, i)).SpecialCells(xlCellTypeVisible).Copy
             ws1.Activate
             ws1.Cells(Rows.count, "C").End(xlUp).offset(1, 0).PasteSpecial
         ElseIf Treffer(i) = "FB" Then
             sh1.Activate
             sh1.Range(Cells(2, i), Cells(Zeilenanzahl, i)).SpecialCells(xlCellTypeVisible).Copy
             ws1.Activate
             ws1.Cells(Rows.count, "A").End(xlUp).offset(1, 0).PasteSpecial
    'AA Knoten kopieren
         ElseIf Treffer(i) = "Objekt-Name" Then
             sh1.Activate
             sh1.Range(Cells(2, i), Cells(Zeilenanzahl, i)).SpecialCells(xlCellTypeVisible).Copy
             ws1.Activate
             ws1.Cells(Rows.count, "F").End(xlUp).offset(1, 0).PasteSpecial
         ElseIf Treffer(i) = "Kommunalität" Then
             sh1.Activate
             sh1.Range(Cells(2, i), Cells(Zeilenanzahl, i)).SpecialCells(xlCellTypeVisible).Copy
             ws1.Activate
             ws1.Cells(Rows.count, "D").End(xlUp).offset(1, 0).PasteSpecial
             
         ElseIf Treffer(i) = "Komponente" Then
             sh1.Activate
             sh1.Range(Cells(2, i), Cells(Zeilenanzahl, i)).SpecialCells(xlCellTypeVisible).Copy
             ws1.Activate
             ws1.Cells(Rows.count, "E").End(xlUp).offset(1, 0).PasteSpecial
         End If
     Next i
            
    ThisWorkbook.Worksheets("Typschl").Activate
    Set RangeStart = ThisWorkbook.Worksheets("Typschl").UsedRange.Find(der, LookIn:=xlValues, Lookat:=xlWhole)
    Typzeile = RangeStart.row
    Gültigkeitsdatum = ThisWorkbook.Worksheets("Typschl").Cells(Typzeile, 7)
    
    If Gültigkeitsdatum = "" Then
        Gültigkeitsdatum = "01.01.9999"
    End If
    
    ThisWorkbook.Worksheets("Typschl").FilterMode = False
                            
    wk1.Close savechanges:=False
    
    ws1.Activate
    Zeilenanzahl = ws1.Cells(Rows.count, 4).End(xlUp).row


    'SA Treffer löschen
    For i = 2 To Zeilenanzahl
        Do While ws1.Cells(i, 4) = "gSA" Or ws1.Cells(i, 4) = "sSA" Or ws1.Cells(i, 4) = "nSA"
            ws1.Rows(i).Delete
            Zeilenanzahl = Zeilenanzahl - 1
        Loop
    Next i
    
    
    'PPG-String kürzen auf KoGr
    For i = 2 To Zeilenanzahl
        Kürzen1 = Left(ws1.Cells(i, 3), 5)
        Kürzen2 = Right(Kürzen1, 4)
        ws1.Cells(i, 3) = Kürzen2
    Next i

    
    'Löschen nicht darstellbarer Zeilen
    For i = 2 To Zeilenanzahl
        Do While ws1.Cells(i, 4) = "x"
            ws1.Rows(i).Delete
            Zeilenanzahl = Zeilenanzahl - 1
            If ws1.Cells(i, 4) = "" Then
                Exit Do
            End If
        Loop
    Next i
    
    
    'Löschen Kleinteile / Formteile Modul KA
    For i = 2 To Zeilenanzahl
        Do While ws1.Cells(i, 2) = "KA"
            ws1.Rows(i).Delete
            Zeilenanzahl = Zeilenanzahl - 1
            If ws1.Cells(i, 2) = "" Then
                Exit Do
            End If
        Loop
    Next i
    
    'EA Eintragen für Zeilen ohne ModulOrg
    For i = 2 To Zeilenanzahl
        Do While ws1.Cells(i, 1) = "" And ws1.Cells(i, 2) = ""
            ws1.Cells(i, 1) = "EA"
        Loop
    Next i

            
    'Sortieren der Spalte FB
    ws1.AutoFilterMode = False
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("VIS_MAKRO").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("VIS_MAKRO").AutoFilter.Sort.SortFields.add Key:=Range( _
        "A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal

    With ActiveWorkbook.Worksheets("VIS_MAKRO").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Pivot-Tabelle erstellen
    Daten = "VIS_MAKRO!R1C1:R" & Zeilenanzahl & "C5"
    ActiveWorkbook.PivotCaches.create(SourceType:=xlDatabase, sourcedata:= _
        Daten, Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:="VIS_PIVOT!R1C1", TableName:="PivotTable", _
        DefaultVersion:=xlPivotTableVersion15
    ws2.Activate
    With ActiveSheet.PivotTables("PivotTable").PivotFields("FB")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable").PivotFields("ModulOrg")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    With ActiveSheet.PivotTables("PivotTable").PivotFields("KoGr")
        .Orientation = xlRowField
        .Position = 3
    End With
    
    With ActiveSheet.PivotTables("PivotTable").PivotFields("Komponente")
        .Orientation = xlRowField
        .Position = 4
    End With
    
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("Treffer"), "Anzahl von Treffer", xlCount
    With ActiveSheet.PivotTables("PivotTable").PivotFields("Treffer")
        .Orientation = xlColumnField
        .Position = 1
    End With
   
    ActiveSheet.PivotTables("PivotTable").PivotFields("FB").AutoSort xlAscending, "FB"
    
   'Pivot-Tabelle auswerten
    Zeilenanzahl = ws2.Cells(Rows.count, 1).End(xlUp).row
    
    For i = 3 To Zeilenanzahl
        If IsNumeric(ws2.Cells(i, 1)) Or ws2.Cells(i, 1) = "(Leer)" Then
        
            Select Case IsEmpty(ws2.Cells(i, 2))
                Case Is = True
                    g = 0
                Case Is = False
                    g = ws2.Cells(i, 2)
            End Select
            
            Select Case IsEmpty(ws2.Cells(i, 3))
                Case Is = True
                    n = 0
                Case Is = False
                    n = ws2.Cells(i, 3)
            End Select
            
            Select Case IsEmpty(ws2.Cells(i, 4))
                Case Is = True
                    s = 0
                Case Is = False
                    s = ws2.Cells(i, 4)
            End Select
            
            If g + s > n Then
                If g > s Then
                    ws2.Cells(i, 7) = "g"
                Else
                    ws2.Cells(i, 7) = "s"
                End If
            Else
                If n = s Then
                    ws2.Cells(i, 7) = "s"
                Else
                    ws2.Cells(i, 7) = "n"
                End If
            End If
                        
            If g > 0 And s > 0 And n > 0 Then
                ws2.Cells(i, 7).Interior.ColorIndex = 16
            ElseIf g > 0 And s > 0 Then
                ws2.Cells(i, 7).Interior.ColorIndex = 16
            ElseIf g > 0 And n > 0 Then
                ws2.Cells(i, 7).Interior.ColorIndex = 16
            ElseIf n > 0 And s > 0 Then
                ws2.Cells(i, 7).Interior.ColorIndex = 16
            End If
        End If
    Next i
    
    For i = 3 To Zeilenanzahl
        j = 1
        If ws2.Cells(i, 7).Interior.ColorIndex = 16 Then
            Do While ws2.Cells(j + i, 7) = "" And j + i < Zeilenanzahl + 1
                ws2.Cells(j + i, 7).Interior.ColorIndex = 16
                j = j + 1
            Loop
        End If
    Next i
    
    'KoGr in Spalte 24 übertragen und forlaufend durchschreiben
    For i = 4 To Zeilenanzahl
        If Len(ws2.Cells(i, 1)) = 4 And IsNumeric(ws2.Cells(i, 1).Value) = False Then
            ws2.Cells(i, 24) = ws2.Cells(i, 1)
        End If
    Next i
    
    For i = 3 To Zeilenanzahl
        j = 1
        If ws2.Cells(i, 24) <> "" Then
            current_value = ws2.Cells(i, 24)
            Do While ws2.Cells(i + j, 24) = ""
                ws2.Cells(i + j, 24) = ws2.Cells(i, 24)
                j = j + 1
                If i + j > Zeilenanzahl Then
                    Exit Do
                End If
            Loop
        End If
    Next i
    
    'Formatieren
    ws2.Cells(1, 7) = "Regel"
    ws2.Cells(1, 7).HorizontalAlignment = xlCenter
    ws2.Cells(1, 7).VerticalAlignment = xlCenter
    ws2.Cells(1, 7).Font.Bold = True
    

    'Fehler-Prozent berechnen
    Richtig = 0
    Falsch = 0
    For i = 4 To Zeilenanzahl
        If ws2.Cells(i, 7).Interior.ColorIndex <> 16 And ws2.Cells(i, 7) <> "" Then
            Richtig = Richtig + ws2.Cells(i, 6)
        ElseIf ws2.Cells(i, 7).Interior.ColorIndex = 16 And ws2.Cells(i, 7) <> "" Then
            g = ws2.Cells(i, 2)
            n = ws2.Cells(i, 3)
            s = ws2.Cells(i, 4)
            If g > n And g > s Then
                Richtig = Richtig + g
                Falsch = Falsch + n + s
            ElseIf s > n And s >= g Then
                Richtig = Richtig + s
                Falsch = Falsch + g + n
            ElseIf n >= g And n >= s Then
                Richtig = Richtig + n
                Falsch = Falsch + s + g
            End If
        End If
    Next i
    
    Anpassung = False
    If MsgBox(Left((Falsch * 100) / (Richtig + Falsch), 4) & "% der Komponenten werden falsch dargestellt!" & vbNewLine & "Möchten Sie diese händisch überarbeiten?", vbYesNoCancel) = vbYes Then
    Anpassung = True
       MsgBox ("Die falsch dargestellten KoGr sind in Spalte ""G"" grau hinterlegt!")
        ws2.Cells(1, 23) = Derivat
        ws2.Cells(2, 23) = Gültigkeitsdatum
        ws2.Cells(3, 23) = Typschlüssel
        'Formatieren
        ActiveSheet.PivotTables("PivotTable").TableStyle2 = "PivotStyleMedium2"
        ws2.Columns("B:D").ColumnWidth = 4
        For i = 4 To Zeilenanzahl
            If ws2.Cells(i, 2) <> "" Then
                ws2.Cells(i, 2).Interior.ColorIndex = 4
            End If
            If ws2.Cells(i, 3) <> "" Then
                ws2.Cells(i, 3).Interior.ColorIndex = 3
            End If
            If ws2.Cells(i, 4) <> "" Then
                ws2.Cells(i, 4).Interior.ColorIndex = 6
            End If
        Next i
        
        ws2.Range(Cells(3, 2), Cells(Zeilenanzahl, 4)).HorizontalAlignment = xlCenter
        ws2.Range(Cells(3, 2), Cells(Zeilenanzahl, 4)).VerticalAlignment = xlCenter
        ws2.Range(Cells(1, 7), Cells(Zeilenanzahl, 7)).AutoFilter
        
        ws2.Activate
    Else
        Call Dateierstellung2(der, Gültigkeitsdatum)
        Call Löschen
        For Each Sheet In ActiveWorkbook.Worksheets
            If Sheet.Name <> "Home" Then
                Sheet.Visible = False
            End If
        Next Sheet
    End If
       
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    
    Call go
    If Anpassung = False Then
    Call Shell("Explorer /e, " & strFile1, vbNormalFocus)
'    MsgBox ("XML-Daten wurden erstellt und im Ordner \Visualisierung_" & Derivat & " gespeichert.")
    End If
                        
End Sub

Sub Dateierstellung2(Derivat As String, Gültigkeitsdatum As String)


    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim strFile2 As String
    Dim Text As String
    Dim Datum As String
    Dim Uhrzeit As String
    Dim Treffer() As String
    Dim Daten As String
    Dim Fachbereich As String
    Dim Eintrag As String
    Dim Grenze(0 To 6) As Integer
    Dim Position As Integer
    Dim Position2 As Integer
    Dim Spaltenzahl As Integer
    Dim Zeilenanzahl As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim g As Integer
    Dim s As Integer
    Dim n As Integer
    Dim o As Integer
    
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    Set ws1 = ThisWorkbook.Worksheets("VIS_MAKRO")
    Set ws2 = ThisWorkbook.Worksheets("VIS_PIVOT")
    
    Datum = Date
    Uhrzeit = Time
    Zeilenanzahl = ws2.Cells(Rows.count, 1).End(xlUp).row
    
        'Grenze bestimmen
    For i = 1 To Zeilenanzahl
        j = 0
        Select Case ws2.Cells(i, 1)
            Case "EA"
                Grenze(0) = i
            Case "EE"
                Grenze(1) = i
            Case "EF"
                Grenze(2) = i
            Case "EI"
                Grenze(3) = i
            Case "EK"
                Grenze(4) = i
            Case "EV"
                Grenze(5) = i
            Case "(Leer)"
                Grenze(6) = i
        End Select
    Next i
    
    'Leere Grenzen eliminieren
    For i = 0 To 6
        If Grenze(i) = 0 Then
            On Error Resume Next
            Grenze(i) = Grenze(i + 1)
            On Error GoTo 0
        End If
    Next i

  
    '*****************************************************************************************************************************************************************************************
    'Farbcode von EI und EK schon auf grün/gelb/rot angepasst
    
    
    'Erstellen der XML-Datei für EA
    strFile2 = strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EA" & ".xml"
    Open strFile2 For Output As #1

 
    'XML-Header für EA
    Print #1, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #1, "<VisualReport xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""VisualReportSchema"" version=""1.0"" author=""Teamcenter Visualization 11.2.2"" date=""" & Datum & """ Time = """ & Uhrzeit & """ >"
        Print #1, "<ReportProp name=""KommVis_EA"" actionType=""changeAppearance"" targetParts=""visible""/>"


    'Regeln für EA
        g = 0
        n = 0
        s = 0
        
        For i = Grenze(0) To Grenze(1) - 1
            If ws2.Cells(i, 7) = "g" Then
                g = g + 1
                Print #1, "<Rule name=""GT" & g & """>"
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #1, "<Condition operator= ""and"">"
                        Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #1, "</Condition>"
                        Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #1, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #1, "</Condition>"
                    Print #1, "</Condition>"
                    Print #1, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #1, "<SimpleClassifier name=""Aktion"">"
                        Print #1, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #1, "<BasicMaterial diffuse=""0.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.000000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #1, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</EnvMap>"
                            Print #1, "<BumpMap>"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</BumpMap>"
                        Print #1, "</Material>"
                    Print #1, "</SimpleClassifier>"
                    Print #1, "</Action>"
                Print #1, "</Rule>"
            ElseIf ws2.Cells(i, 7) = "n" Then
                n = n + 1
                Print #1, "<Rule name=""NT" & n & """>"
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #1, "<Condition operator= ""and"">"
                        Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #1, "</Condition>"
                        Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #1, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #1, "</Condition>"
                    Print #1, "</Condition>"
                    Print #1, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #1, "<SimpleClassifier name=""Aktion"">"
                        Print #1, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #1, "<BasicMaterial diffuse=""1.000000 0.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.000000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #1, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</EnvMap>"
                            Print #1, "<BumpMap>"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</BumpMap>"
                        Print #1, "</Material>"
                    Print #1, "</SimpleClassifier>"
                    Print #1, "</Action>"
                Print #1, "</Rule>"
            ElseIf ws2.Cells(i, 7) = "s" Then
                s = s + 1
                Print #1, "<Rule name=""ST" & s & """>"
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #1, "<Condition operator= ""and"">"
                        Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #1, "</Condition>"
                        Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #1, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #1, "</Condition>"
                    Print #1, "</Condition>"
                    Print #1, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #1, "<SimpleClassifier name=""Aktion"">"
                        Print #1, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #1, "<BasicMaterial diffuse=""1.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #1, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                            Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</EnvMap>"
                            Print #1, "<BumpMap>"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</BumpMap>"
                        Print #1, "</Material>"
                    Print #1, "</SimpleClassifier>"
                    Print #1, "</Action>"
                Print #1, "</Rule>"
              End If
        Next i
        
        'AA hier weiter
        '(Leer)-Komponenten aus dem Motor
        If ws2.Cells(Grenze(1), 7) = "g" Then
            o = 1
            g = g + 1
            
            
            Print #1, "<Rule name=""GT" & g & """>"
                    
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #1, "<Condition operator= ""or"">"
                    
                    Do While o <= 15
                        ModulZahl = Format(o, "00")
                        Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MA" & Format(o, "00") & """ type= ""attribute"">"
                            Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #1, "</Condition>"
                        o = o + 1
                    Loop
                    o = 1
                    Do While o <= 4
                        ModulZahl = Format(o, "00")
                        Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MB" & Format(o, "00") & """ type= ""attribute"">"
                            Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #1, "</Condition>"
                        o = o + 1
                    Loop
                    o = 1
                    Do While o <= 8
                        ModulZahl = Format(o, "00")
                        Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MC" & Format(o, "00") & """ type= ""attribute"">"
                            Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #1, "</Condition>"
                        o = o + 1
                    Loop
                    o = 1
                    Do While o <= 7
                        ModulZahl = Format(o, "00")
                        Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MD" & Format(o, "00") & """ type= ""attribute"">"
                            Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #1, "</Condition>"
                        o = o + 1
                    Loop
                    
                    Print #1, "</Condition>"
                    Print #1, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #1, "<SimpleClassifier name=""Aktion"">"
                        Print #1, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #1, "<BasicMaterial diffuse=""0.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.000000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #1, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</EnvMap>"
                            Print #1, "<BumpMap>"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</BumpMap>"
                        Print #1, "</Material>"
                    Print #1, "</SimpleClassifier>"
                    Print #1, "</Action>"
                Print #1, "</Rule>"
                
        ElseIf ws2.Cells(Grenze(1), 7) = "n" Then
            o = 1
            n = n + 1
            Print #1, "<Rule name=""NT" & n & """>"
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #1, "<Condition operator= ""or"">"
                    
                    Do While o <= 15
                        
                        Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MA" & Format(o, "00") & """ type= ""attribute"">"
                            Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #1, "</Condition>"
                        o = o + 1
                    Loop
                    o = 1
                    Do While o <= 4
                        
                        Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MB" & Format(o, "00") & """ type= ""attribute"">"
                            Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #1, "</Condition>"
                        o = o + 1
                    Loop
                    o = 1
                    Do While o <= 8
                        
                        Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MC" & Format(o, "00") & """ type= ""attribute"">"
                            Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #1, "</Condition>"
                        o = o + 1
                    Loop
                    o = 1
                    Do While o <= 7
                        
                        Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MD" & Format(o, "00") & """ type= ""attribute"">"
                            Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #1, "</Condition>"
                        o = o + 1
                    Loop
                    
                    Print #1, "</Condition>"
                    Print #1, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #1, "<SimpleClassifier name=""Aktion"">"
                        Print #1, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #1, "<BasicMaterial diffuse=""1.000000 0.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.000000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #1, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</EnvMap>"
                            Print #1, "<BumpMap>"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</BumpMap>"
                        Print #1, "</Material>"
                    Print #1, "</SimpleClassifier>"
                    Print #1, "</Action>"
                Print #1, "</Rule>"
        ElseIf ws2.Cells(Grenze(1), 7) = "s" Then
            o = 1
            s = s + 1
                Print #1, "<Rule name=""ST" & s & """>"
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #1, "<Condition operator= ""or"">"
                    
                    Do While o <= 15
                    
                    Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MA" & Format(o, "00") & """ type= ""attribute"">"
                        Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                    Print #1, "</Condition>"
                    o = o + 1
                    Loop
                    o = 1
                    Do While o <= 4
                    Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MB" & Format(o, "00") & """ type= ""attribute"">"
                        Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                    Print #1, "</Condition>"
                    o = o + 1
                    Loop
                    o = 1
                    Do While o <= 8
                    Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MC" & Format(o, "00") & """ type= ""attribute"">"
                        Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                    Print #1, "</Condition>"
                    o = o + 1
                    Loop
                    o = 1
                    Do While o <= 7
                    Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MD" & Format(o, "00") & """ type= ""attribute"">"
                        Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                    Print #1, "</Condition>"
                    o = o + 1
                    Loop
                        
                    Print #1, "</Condition>"
                    Print #1, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #1, "<SimpleClassifier name=""Aktion"">"
                        Print #1, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #1, "<BasicMaterial diffuse=""1.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #1, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</Texture>"
                            Print #1, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                            Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</EnvMap>"
                            Print #1, "<BumpMap>"
                                Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #1, "</BumpMap>"
                        Print #1, "</Material>"
                    Print #1, "</SimpleClassifier>"
                    Print #1, "</Action>"
                Print #1, "</Rule>"
        End If

    'Regel für "Keine Aussage"
        Print #1, "<Action type=""nonMatched"" displayMode=""solid wireframe"">"
            Print #1, "<SimpleClassifier>"
                Print #1, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                    Print #1, "<BasicMaterial diffuse=""0.498039 0.498039 0.498039"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.231588 0.231588 0.231588"" transparency=""0.750000"" shininess=""0.300000""/>"
                    Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                        Print #1, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                    Print #1, "</Texture>"
                    Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                        Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #1, "</Texture>"
                    Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                        Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #1, "</Texture>"
                    Print #1, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                        Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #1, "</Texture>"
                    Print #1, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                        Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #1, "</EnvMap>"
                    Print #1, "<BumpMap>"
                        Print #1, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #1, "</BumpMap>"
                Print #1, "</Material>"
            Print #1, "</SimpleClassifier>"
        Print #1, "</Action>"
    Print #1, "</VisualReport>"

    Close #1
    
    
    
    
    '######################################################
    'Probleme mit dem umbennen der Datei
    Name strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EA" & ".xml" As strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EA" & ".vpx"
    
    
'**************************************************************************************************************************************************************************************************************************************************************************************************************************
    
    'Erstellen der XML-Datei für EE
    strFile2 = strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EE" & ".xml"
    Open strFile2 For Output As #2

 
    'XML-Header für EE
    Print #2, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #2, "<VisualReport xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""VisualReportSchema"" version=""1.0"" author=""Teamcenter Visualization 11.2.2"" date=""" & Datum & """ Time = """ & Uhrzeit & """ >"
        Print #2, "<ReportProp name=""ModOrg_Filter"" actionType=""changeAppearance"" targetParts=""visible""/>"


    'Regeln für EE
        g = 0
        n = 0
        s = 0
        
        For i = Grenze(1) + 2 To Grenze(2) - 1
            If ws2.Cells(i, 7) = "g" Then
                g = g + 1
                Print #2, "<Rule name=""GT" & g & """>"
                    Print #2, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #2, "<Condition operator= ""and"">"
                        Print #2, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #2, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #2, "</Condition>"
                        Print #2, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #2, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #2, "</Condition>"
                    Print #2, "</Condition>"
                    Print #2, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #2, "<SimpleClassifier name=""Aktion"">"
                        Print #2, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #2, "<BasicMaterial diffuse=""0.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.000000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #2, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #2, "</Texture>"
                            Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</Texture>"
                            Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</Texture>"
                            Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</Texture>"
                            Print #2, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                                Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</EnvMap>"
                            Print #2, "<BumpMap>"
                                Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</BumpMap>"
                        Print #2, "</Material>"
                    Print #2, "</SimpleClassifier>"
                    Print #2, "</Action>"
                Print #2, "</Rule>"
            ElseIf ws2.Cells(i, 7) = "n" Then
                n = n + 1
                Print #2, "<Rule name=""NT" & n & """>"
                    Print #2, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #2, "<Condition operator= ""and"">"
                        Print #2, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #2, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #2, "</Condition>"
                        Print #2, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #2, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #2, "</Condition>"
                    Print #2, "</Condition>"
                    Print #2, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #2, "<SimpleClassifier name=""Aktion"">"
                        Print #2, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #2, "<BasicMaterial diffuse=""1.000000 0.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.000000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #2, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #2, "</Texture>"
                            Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</Texture>"
                            Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</Texture>"
                            Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</Texture>"
                            Print #2, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                                Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</EnvMap>"
                            Print #2, "<BumpMap>"
                                Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</BumpMap>"
                        Print #2, "</Material>"
                    Print #2, "</SimpleClassifier>"
                    Print #2, "</Action>"
                Print #2, "</Rule>"
            ElseIf ws2.Cells(i, 7) = "s" Then
                s = s + 1
                Print #2, "<Rule name=""ST" & s & """>"
                    Print #2, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #2, "<Condition operator= ""and"">"
                        Print #2, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #2, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #2, "</Condition>"
                        Print #2, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #2, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #2, "</Condition>"
                    Print #2, "</Condition>"
                    Print #2, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #2, "<SimpleClassifier name=""Aktion"">"
                        Print #2, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #2, "<BasicMaterial diffuse=""1.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #2, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #2, "</Texture>"
                            Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</Texture>"
                            Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</Texture>"
                            Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</Texture>"
                            Print #2, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                            Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</EnvMap>"
                            Print #2, "<BumpMap>"
                                Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #2, "</BumpMap>"
                        Print #2, "</Material>"
                    Print #2, "</SimpleClassifier>"
                    Print #2, "</Action>"
                Print #2, "</Rule>"
              End If
        Next i
        
    'Regel für "Keine Aussage"
        Print #2, "<Action type=""nonMatched"" displayMode=""solid wireframe"">"
            Print #2, "<SimpleClassifier>"
                Print #2, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                    Print #2, "<BasicMaterial diffuse=""0.498039 0.498039 0.498039"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.231588 0.231588 0.231588"" transparency=""0.750000"" shininess=""0.300000""/>"
                    Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                        Print #2, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                    Print #2, "</Texture>"
                    Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                        Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #2, "</Texture>"
                    Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                        Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #2, "</Texture>"
                    Print #2, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                        Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #2, "</Texture>"
                    Print #2, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                        Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #2, "</EnvMap>"
                    Print #2, "<BumpMap>"
                        Print #2, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #2, "</BumpMap>"
                Print #2, "</Material>"
            Print #2, "</SimpleClassifier>"
        Print #2, "</Action>"
    Print #2, "</VisualReport>"

    Close #2
    
    
    'Rename
    Name strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EE" & ".xml" As strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EE" & ".vpx"
    
'*****************************************************************************************************************************************************************************************************************************************************************************************************
    
    'Erstellen der XML-Datei für EF
    strFile2 = strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EF" & ".xml"
    Open strFile2 For Output As #3

 
    'XML-Header für EF
    Print #3, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #3, "<VisualReport xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""VisualReportSchema"" version=""1.0"" author=""Teamcenter Visualization 11.2.2"" date=""" & Datum & """ Time = """ & Uhrzeit & """ >"
        Print #3, "<ReportProp name=""KommVis_EF"" actionType=""changeAppearance"" targetParts=""visible""/>"


    'Regeln für EF
        g = 0
        n = 0
        s = 0
        
        For i = Grenze(2) + 2 To Grenze(3) - 1
            If ws2.Cells(i, 7) = "g" Then
                g = g + 1
                Print #3, "<Rule name=""GT" & g & """>"
                    Print #3, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #3, "<Condition operator= ""and"">"
                        Print #3, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #3, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #3, "</Condition>"
                        Print #3, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #3, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #3, "</Condition>"
                    Print #3, "</Condition>"
                    Print #3, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #3, "<SimpleClassifier name=""Aktion"">"
                        Print #3, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #3, "<BasicMaterial diffuse=""0.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.000000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #3, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #3, "</Texture>"
                            Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</Texture>"
                            Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</Texture>"
                            Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</Texture>"
                            Print #3, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                                Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</EnvMap>"
                            Print #3, "<BumpMap>"
                                Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</BumpMap>"
                        Print #3, "</Material>"
                    Print #3, "</SimpleClassifier>"
                    Print #3, "</Action>"
                Print #3, "</Rule>"
            ElseIf ws2.Cells(i, 7) = "n" Then
                n = n + 1
                Print #3, "<Rule name=""NT" & n & """>"
                    Print #3, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #3, "<Condition operator= ""and"">"
                        Print #3, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #3, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #3, "</Condition>"
                        Print #3, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #3, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #3, "</Condition>"
                    Print #3, "</Condition>"
                    Print #3, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #3, "<SimpleClassifier name=""Aktion"">"
                        Print #3, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #3, "<BasicMaterial diffuse=""1.000000 0.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.000000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #3, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #3, "</Texture>"
                            Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</Texture>"
                            Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</Texture>"
                            Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</Texture>"
                            Print #3, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                                Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</EnvMap>"
                            Print #3, "<BumpMap>"
                                Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</BumpMap>"
                        Print #3, "</Material>"
                    Print #3, "</SimpleClassifier>"
                    Print #3, "</Action>"
                Print #3, "</Rule>"
            ElseIf ws2.Cells(i, 7) = "s" Then
                s = s + 1
                Print #3, "<Rule name=""ST" & s & """>"
                    Print #3, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #3, "<Condition operator= ""and"">"
                        Print #3, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #3, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #3, "</Condition>"
                        Print #3, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #3, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #3, "</Condition>"
                    Print #3, "</Condition>"
                    Print #3, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #3, "<SimpleClassifier name=""Aktion"">"
                        Print #3, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #3, "<BasicMaterial diffuse=""1.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #3, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #3, "</Texture>"
                            Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</Texture>"
                            Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</Texture>"
                            Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</Texture>"
                            Print #3, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                            Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</EnvMap>"
                            Print #3, "<BumpMap>"
                                Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #3, "</BumpMap>"
                        Print #3, "</Material>"
                    Print #3, "</SimpleClassifier>"
                    Print #3, "</Action>"
                Print #3, "</Rule>"
              End If
        Next i
        
    'Regel für "Keine Aussage"
        Print #3, "<Action type=""nonMatched"" displayMode=""solid wireframe"">"
            Print #3, "<SimpleClassifier>"
                Print #3, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                    Print #3, "<BasicMaterial diffuse=""0.498039 0.498039 0.498039"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.231588 0.231588 0.231588"" transparency=""0.750000"" shininess=""0.300000""/>"
                    Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                        Print #3, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                    Print #3, "</Texture>"
                    Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                        Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #3, "</Texture>"
                    Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                        Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #3, "</Texture>"
                    Print #3, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                        Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #3, "</Texture>"
                    Print #3, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                        Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #3, "</EnvMap>"
                    Print #3, "<BumpMap>"
                        Print #3, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #3, "</BumpMap>"
                Print #3, "</Material>"
            Print #3, "</SimpleClassifier>"
        Print #3, "</Action>"
    Print #3, "</VisualReport>"

    Close #3
    
    
    'Rename
    Name strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EF" & ".xml" As strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EF" & ".vpx"
    
'*****************************************************************************************************************************************************************************************************************************************************************************************************
    
    
   'Erstellen der XML-Datei für EI
    strFile2 = strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EI" & ".xml"
    Open strFile2 For Output As #4

 
    'XML-Header für EI
    Print #4, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #4, "<VisualReport xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""VisualReportSchema"" version=""1.0"" author=""Teamcenter Visualization 11.2.2"" date=""" & Datum & """ Time = """ & Uhrzeit & """ >"
        Print #4, "<ReportProp name=""KommVis_EI"" actionType=""changeAppearance"" targetParts=""visible""/>"


    'Regeln für EI
        g = 0
        n = 0
        s = 0
        
        For i = Grenze(3) + 2 To Grenze(4) - 1
            If ws2.Cells(i, 7) = "g" Then
                g = g + 1
                Print #4, "<Rule name=""GT" & g & """>"
                    Print #4, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #4, "<Condition operator= ""and"">"
                        Print #4, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #4, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #4, "</Condition>"
                        Print #4, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #4, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #4, "</Condition>"
                    Print #4, "</Condition>"
                    Print #4, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #4, "<SimpleClassifier name=""Aktion"">"
                        Print #4, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #4, "<BasicMaterial diffuse=""0.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.000000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #4, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #4, "</Texture>"
                            Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</Texture>"
                            Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</Texture>"
                            Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</Texture>"
                            Print #4, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                                Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</EnvMap>"
                            Print #4, "<BumpMap>"
                                Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</BumpMap>"
                        Print #4, "</Material>"
                    Print #4, "</SimpleClassifier>"
                    Print #4, "</Action>"
                Print #4, "</Rule>"
            ElseIf ws2.Cells(i, 7) = "n" Then
                n = n + 1
                Print #4, "<Rule name=""NT" & n & """>"
                    Print #4, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #4, "<Condition operator= ""and"">"
                        Print #4, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #4, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #4, "</Condition>"
                        Print #4, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #4, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #4, "</Condition>"
                    Print #4, "</Condition>"
                    Print #4, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #4, "<SimpleClassifier name=""Aktion"">"
                        Print #4, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #4, "<BasicMaterial diffuse=""1.000000 0.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.000000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #4, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #4, "</Texture>"
                            Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</Texture>"
                            Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</Texture>"
                            Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</Texture>"
                            Print #4, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                                Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</EnvMap>"
                            Print #4, "<BumpMap>"
                                Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</BumpMap>"
                        Print #4, "</Material>"
                    Print #4, "</SimpleClassifier>"
                    Print #4, "</Action>"
                Print #4, "</Rule>"
            ElseIf ws2.Cells(i, 7) = "s" Then
                s = s + 1
                Print #4, "<Rule name=""ST" & s & """>"
                    Print #4, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #4, "<Condition operator= ""and"">"
                        Print #4, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #4, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #4, "</Condition>"
                        Print #4, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #4, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #4, "</Condition>"
                    Print #4, "</Condition>"
                    Print #4, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #4, "<SimpleClassifier name=""Aktion"">"
                        Print #4, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #4, "<BasicMaterial diffuse=""1.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #4, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #4, "</Texture>"
                            Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</Texture>"
                            Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</Texture>"
                            Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</Texture>"
                            Print #4, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                            Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</EnvMap>"
                            Print #4, "<BumpMap>"
                                Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #4, "</BumpMap>"
                        Print #4, "</Material>"
                    Print #4, "</SimpleClassifier>"
                    Print #4, "</Action>"
                Print #4, "</Rule>"
              End If
        Next i
        
    'Regel für "Keine Aussage"
        Print #4, "<Action type=""nonMatched"" displayMode=""solid wireframe"">"
            Print #4, "<SimpleClassifier>"
                Print #4, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                    Print #4, "<BasicMaterial diffuse=""0.498039 0.498039 0.498039"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.231588 0.231588 0.231588"" transparency=""0.750000"" shininess=""0.300000""/>"
                    Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                        Print #4, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                    Print #4, "</Texture>"
                    Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                        Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #4, "</Texture>"
                    Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                        Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #4, "</Texture>"
                    Print #4, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                        Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #4, "</Texture>"
                    Print #4, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                        Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #4, "</EnvMap>"
                    Print #4, "<BumpMap>"
                        Print #4, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #4, "</BumpMap>"
                Print #4, "</Material>"
            Print #4, "</SimpleClassifier>"
        Print #4, "</Action>"
    Print #4, "</VisualReport>"

    Close #4
    
    
    'Rename
    Name strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EI" & ".xml" As strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EI" & ".vpx"
    
'*****************************************************************************************************************************************************************************************************************************************************************************************************
     
    
    'Erstellen der XML-Datei für EK
    strFile2 = strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EK" & ".xml"
    Open strFile2 For Output As #5

 
    'XML-Header für EK
    Print #5, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #5, "<VisualReport xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""VisualReportSchema"" version=""1.0"" author=""Teamcenter Visualization 11.2.2"" date=""" & Datum & """ Time = """ & Uhrzeit & """ >"
        Print #5, "<ReportProp name=""KommVis_EK"" actionType=""changeAppearance"" targetParts=""visible""/>"


    'Regeln für EK
        g = 0
        n = 0
        s = 0
        
        For i = Grenze(4) + 2 To Grenze(5) - 1
            If ws2.Cells(i, 7) = "g" Then
                g = g + 1
                Print #5, "<Rule name=""GT" & g & """>"
                    Print #5, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #5, "<Condition operator= ""and"">"
                        Print #5, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #5, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #5, "</Condition>"
                        Print #5, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #5, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #5, "</Condition>"
                    Print #5, "</Condition>"
                    Print #5, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #5, "<SimpleClassifier name=""Aktion"">"
                        Print #5, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #5, "<BasicMaterial diffuse=""0.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.000000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #5, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #5, "</Texture>"
                            Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</Texture>"
                            Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</Texture>"
                            Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</Texture>"
                            Print #5, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                                Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</EnvMap>"
                            Print #5, "<BumpMap>"
                                Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</BumpMap>"
                        Print #5, "</Material>"
                    Print #5, "</SimpleClassifier>"
                    Print #5, "</Action>"
                Print #5, "</Rule>"
            ElseIf ws2.Cells(i, 7) = "n" Then
                n = n + 1
                Print #5, "<Rule name=""NT" & n & """>"
                    Print #5, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #5, "<Condition operator= ""and"">"
                        Print #5, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #5, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #5, "</Condition>"
                        Print #5, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #5, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #5, "</Condition>"
                    Print #5, "</Condition>"
                    Print #5, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #5, "<SimpleClassifier name=""Aktion"">"
                        Print #5, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #5, "<BasicMaterial diffuse=""1.000000 0.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.000000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #5, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #5, "</Texture>"
                            Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</Texture>"
                            Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</Texture>"
                            Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</Texture>"
                            Print #5, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                                Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</EnvMap>"
                            Print #5, "<BumpMap>"
                                Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</BumpMap>"
                        Print #5, "</Material>"
                    Print #5, "</SimpleClassifier>"
                    Print #5, "</Action>"
                Print #5, "</Rule>"
            ElseIf ws2.Cells(i, 7) = "s" Then
                s = s + 1
                Print #5, "<Rule name=""ST" & s & """>"
                    Print #5, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #5, "<Condition operator= ""and"">"
                        Print #5, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #5, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #5, "</Condition>"
                        Print #5, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #5, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #5, "</Condition>"
                    Print #5, "</Condition>"
                    Print #5, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #5, "<SimpleClassifier name=""Aktion"">"
                        Print #5, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #5, "<BasicMaterial diffuse=""1.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #5, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #5, "</Texture>"
                            Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</Texture>"
                            Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</Texture>"
                            Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</Texture>"
                            Print #5, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                            Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</EnvMap>"
                            Print #5, "<BumpMap>"
                                Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #5, "</BumpMap>"
                        Print #5, "</Material>"
                    Print #5, "</SimpleClassifier>"
                    Print #5, "</Action>"
                Print #5, "</Rule>"
              End If
        Next i
        
    'Regel für "Keine Aussage"
        Print #5, "<Action type=""nonMatched"" displayMode=""solid wireframe"">"
            Print #5, "<SimpleClassifier>"
                Print #5, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                    Print #5, "<BasicMaterial diffuse=""0.498039 0.498039 0.498039"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.231588 0.231588 0.231588"" transparency=""0.750000"" shininess=""0.300000""/>"
                    Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                        Print #5, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                    Print #5, "</Texture>"
                    Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                        Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #5, "</Texture>"
                    Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                        Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #5, "</Texture>"
                    Print #5, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                        Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #5, "</Texture>"
                    Print #5, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                        Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #5, "</EnvMap>"
                    Print #5, "<BumpMap>"
                        Print #5, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #5, "</BumpMap>"
                Print #5, "</Material>"
            Print #5, "</SimpleClassifier>"
        Print #5, "</Action>"
    Print #5, "</VisualReport>"

    Close #5
    
    
    'Rename
    Name strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EK" & ".xml" As strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EK" & ".vpx"
    
'*****************************************************************************************************************************************************************************************************************************************************************************************************
     
    
   'Erstellen der XML-Datei für EV
    strFile2 = strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EV" & ".xml"
    Open strFile2 For Output As #6

 
    'XML-Header für EV
    Print #6, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #6, "<VisualReport xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""VisualReportSchema"" version=""1.0"" author=""Teamcenter Visualization 11.2.2"" date=""" & Datum & """ Time = """ & Uhrzeit & """ >"
        Print #6, "<ReportProp name=""KommVis_EV"" actionType=""changeAppearance"" targetParts=""visible""/>"


    'Regeln für EV
        g = 0
        n = 0
        s = 0
        
        For i = Grenze(5) + 2 To Grenze(6) - 1
            If ws2.Cells(i, 7) = "g" Then
                g = g + 1
                Print #6, "<Rule name=""GT" & g & """>"
                    Print #6, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #6, "<Condition operator= ""and"">"
                        Print #6, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #6, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #6, "</Condition>"
                        Print #6, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #6, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #6, "</Condition>"
                    Print #6, "</Condition>"
                    Print #6, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #6, "<SimpleClassifier name=""Aktion"">"
                        Print #6, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #6, "<BasicMaterial diffuse=""0.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.000000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #6, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #6, "</Texture>"
                            Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</Texture>"
                            Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</Texture>"
                            Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</Texture>"
                            Print #6, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                                Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</EnvMap>"
                            Print #6, "<BumpMap>"
                                Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</BumpMap>"
                        Print #6, "</Material>"
                    Print #6, "</SimpleClassifier>"
                    Print #6, "</Action>"
                Print #6, "</Rule>"
            ElseIf ws2.Cells(i, 7) = "n" Then
                n = n + 1
                Print #6, "<Rule name=""NT" & n & """>"
                    Print #6, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #6, "<Condition operator= ""and"">"
                        Print #6, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #6, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #6, "</Condition>"
                        Print #6, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #6, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #6, "</Condition>"
                    Print #6, "</Condition>"
                    Print #6, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #6, "<SimpleClassifier name=""Aktion"">"
                        Print #6, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #6, "<BasicMaterial diffuse=""1.000000 0.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.000000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #6, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #6, "</Texture>"
                            Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</Texture>"
                            Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</Texture>"
                            Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</Texture>"
                            Print #6, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                                Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</EnvMap>"
                            Print #6, "<BumpMap>"
                                Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</BumpMap>"
                        Print #6, "</Material>"
                    Print #6, "</SimpleClassifier>"
                    Print #6, "</Action>"
                Print #6, "</Rule>"
            ElseIf ws2.Cells(i, 7) = "s" Then
                s = s + 1
                Print #6, "<Rule name=""ST" & s & """>"
                    Print #6, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Print #6, "<Condition operator= ""and"">"
                        Print #6, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 24) & """ type= ""attribute"">"
                            Print #6, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                        Print #6, "</Condition>"
                        Print #6, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & ws2.Cells(i, 1) & """ type= ""attribute"">"
                            Print #6, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
                        Print #6, "</Condition>"
                    Print #6, "</Condition>"
                    Print #6, "<Action type=""matched"" displayMode=""solid wireframe"">"
                    Print #6, "<SimpleClassifier name=""Aktion"">"
                        Print #6, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                            Print #6, "<BasicMaterial diffuse=""1.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
                            Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                                Print #6, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                            Print #6, "</Texture>"
                            Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                                Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</Texture>"
                            Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                                Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</Texture>"
                            Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                                Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</Texture>"
                            Print #6, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                            Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</EnvMap>"
                            Print #6, "<BumpMap>"
                                Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                            Print #6, "</BumpMap>"
                        Print #6, "</Material>"
                    Print #6, "</SimpleClassifier>"
                    Print #6, "</Action>"
                Print #6, "</Rule>"
              End If
        Next i
        
    'Regel für "Keine Aussage"
        Print #6, "<Action type=""nonMatched"" displayMode=""solid wireframe"">"
            Print #6, "<SimpleClassifier>"
                Print #6, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                    Print #6, "<BasicMaterial diffuse=""0.498039 0.498039 0.498039"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.231588 0.231588 0.231588"" transparency=""0.750000"" shininess=""0.300000""/>"
                    Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
                        Print #6, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
                    Print #6, "</Texture>"
                    Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
                        Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #6, "</Texture>"
                    Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
                        Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #6, "</Texture>"
                    Print #6, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
                        Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #6, "</Texture>"
                    Print #6, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
                        Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #6, "</EnvMap>"
                    Print #6, "<BumpMap>"
                        Print #6, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
                    Print #6, "</BumpMap>"
                Print #6, "</Material>"
            Print #6, "</SimpleClassifier>"
        Print #6, "</Action>"
    Print #6, "</VisualReport>"

    Close #6
    
    
    'Rename
    Name strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EV" & ".xml" As strFile1 & "\Visu" & Derivat & "_" & Gültigkeitsdatum & "_" & "EV" & ".vpx"

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    

End Sub

Sub Löschen()

    Dim Zeilenanzahl As Integer
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    
    Set ws1 = ThisWorkbook.Worksheets("VIS_MAKRO")
    Set ws2 = ThisWorkbook.Worksheets("VIS_PIVOT")
    ws1.Activate
    
    Zeilenanzahl = ws1.Cells(Rows.count, 1).End(xlUp).row
    
    ws1.Activate
    ws1.Range(Cells(2, 1), Cells(Zeilenanzahl + 100, 6)).Select
    Selection.Delete Shift:=xlUp
    
    ws2.Activate
    Zeilenanzahl = ws2.Cells(Rows.count, 1).End(xlUp).row
    ws2.Range(Cells(1, 24), Cells(Zeilenanzahl + 100, 24)).Select
    Selection.Delete Shift:=xlUp
    ws2.Range(Cells(1, 1), Cells(Zeilenanzahl + 100, 7)).Select
    Selection.Delete Shift:=xlUp
    ws2.Range(Cells(1, 23), Cells(Zeilenanzahl + 100, 23)).Select
    Selection.Delete Shift:=xlUp
    If ws2.Cells(1, 1) <> "" Then
        ws2.Range(Cells(1, 7), Cells(Zeilenanzahl, 7)).AutoFilter
    End If

End Sub










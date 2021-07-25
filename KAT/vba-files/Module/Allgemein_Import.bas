Attribute VB_Name = "Allgemein_Import"
Option Explicit

Sub importnewdata()

    Dim Var
    Dim str As String
    Dim wb As Workbook, wbk As Workbook
    Dim fileType As String
    Dim i As Integer
    Dim success As Boolean, found As Boolean
    Dim sl As SlicerCache
    ' Dim KD As String
    Dim Derivat As String
    Dim strSuchenNach As String
    Dim strErsetzenMit As String
    

    ' checks if megaliste and the folder to store pie charts exist
    Call checkToolResources
    Var = Application.GetOpenFilename(Title:="Please choose a file to open", MultiSelect:=True)
    ' VarType(var) = 11 --> checks if var is boolean or string
    If VarType(Var) = 11 Then Exit Sub ' basically if var = Falsch, which means that the user canceled the procedure or escaped it
    
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    success = False
    For i = 1 To UBound(Var)
        If Right(Var(i), 5) = ".xlsx" Then
            Set wb = Workbooks.Open(Var(i))
            '' analyse the file type to direct to the appropriate response
            '' should the file be treated as a SAP export or a KP file
            fileType = checkFileFormat(wb)
            If fileType = "SAP Export" Then
            'Abfrage ob Erstanläufer, Weite Ableitung oder Engeableitung, dann Auswahl der richtigen KD-Referenz und Übertragung zum Typschlüssel
            strSuchenNach = ".xlsx"
            strErsetzenMit = ""
            Derivat = Replace(wb.Name, strSuchenNach, strErsetzenMit)
            'KD = vbNullString
            'Do While KD = vbNullString Or KD = "Falsch"
            'KD = Application.InputBox("Ist dieses Derivat " & derivat & " ein Erstanläufer (E), eine Enge Ableitung (EA) oder eine Weite Ableitung (WA)?" & vbNewLine _
                              & "Bitte wählen Sie eine der drei Abkürzungen. Beispiel: E oder EA oder WA")
            'Loop
                'Call checkTypschl(wb.Name, KD)
                Call checkTypschl(wb.Name)
                'Kennzahlen
                'Call Kenn(wb.Name, KD)
                Call Kenn(wb.Name)
                Call addSAP(wb.Name)
                success = True
    
            ElseIf fileType = "KP File" Then
                Call addKP(wb)
                success = True
            Else
                MsgBox "The format of the file wasn't recognized: " & vbNewLine & Var(i) & vbNewLine & fileType
            End If
            wb.Close False
        Else
            MsgBox "The file to import must be in the excel format (.xlsx) :" & vbNewLine & Var(i)
        End If
    Next i
    
    If success = True Then
        Call createPivot
        Workbooks("MEGALISTE.xlsx").Sheets("Derivat").Cells.WrapText = False
        Workbooks("MEGALISTE.xlsx").Close True
    Else
        For Each wbk In Workbooks
            '' if megaliste is open then close it without saving
            If wbk.Name = "MEGALISTE.xlsx" Then
                Workbooks("MEGALISTE.xlsx").Close False
                Exit For
            End If
        Next wbk
    End If

    '' display the first slicer (derivat)
    If selectedDerivatCount > 0 Then
        With ThisWorkbook
            freezeTriger = True
            found = False
            For Each sl In ThisWorkbook.SlicerCaches
                If sl.Slicers(1).Name = "Derivat" Then found = True: Exit For
            Next sl
            If found = False Then
                .SlicerCaches.Add2(.Sheets("PIVOT").PivotTables("PivotTableMEGALISTE"), "Derivat").Slicers.add .Sheets("Home"), , "Derivat", "Derivat", 10, 180, 135, 165
            End If
            freezeTriger = False
        End With
    End If
    
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
End Sub
Sub checkToolResources()
    
    '' check if general files are there, the ones very important, whitout which the tool cannot function
    Dim pfad As String

    pfad = ThisWorkbook.Path & "\KAT_Vorlage\MEGALISTE.xlsx"
    
    If Dir(pfad) = vbNullString Then
        MsgBox "Missing MEGALISTE.xlsx in " & ThisWorkbook.Path & "\KAT_Vorlage\"
        End
    End If
    
    If Dir(ThisWorkbook.Path & "\KAT_Vorlage\Heatmap_Chart_Diagramm\", vbDirectory) = vbNullString Then
         MkDir ThisWorkbook.Path & "\KAT_Vorlage\Heatmap_Chart_Diagramm"
    End If
    
    
End Sub


Function checkFileFormat(wb As Workbook) As String

    Dim sh As Worksheet, shKopf As Worksheet
    Dim colName() As Variant
    Dim col() As Variant
    Dim i As Integer

    'check if file has the SAP export format: Kopf mit Parameter, Strukturbericht, column names, Typschlüsselliste
    If wb.Sheets(1).Name = "Kopf mit Parameter" Then
        Set shKopf = wb.Sheets("Kopf mit Parameter")
        
        If wb.Sheets(2).Name <> "Strukturbericht" Then
            checkFileFormat = "Missing 'Struktubericht' table in file."
            Exit Function
        End If
        
        If shKopf.Range("B35") = vbNullString Then
            checkFileFormat = "Missing Gültigkeit Datum in cell B35 in worksheet 'Kopf mit Parameter'."
            Exit Function
        End If
        '' to do : block execution if 4 main columns not found in the header
        checkFileFormat = "SAP Export"
        
    ElseIf InStr(wb.Name, "KP") > 0 Then
        If InStr(wb.Sheets(1).Name, "KP") < 1 Then
            checkFileFormat = "The first sheet in the file does not contain the word 'Konfigurationsprämissen'."
            Exit Function
        End If
        
        Set sh = wb.Sheets(1)
        colName() = Array("Modul", "ModulBezeichnung", "Bezeichnung")
        col() = Array("A", "B", "C")
        For i = 0 To UBound(colName)
            If sh.Range(col(i) & 1).Value <> colName(i) Then
                checkFileFormat = "Cell " & col(i) & "1 in Derivat should contain " & colName(i)
                Exit Function
            End If
        Next i
        checkFileFormat = "KP File"
        
    End If
    
End Function

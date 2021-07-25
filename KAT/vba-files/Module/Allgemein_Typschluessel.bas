Attribute VB_Name = "Allgemein_Typschluessel"
Option Explicit

Sub displayTypschl()

    '' this function allows to hide or dispaly the Typschlüsselliste
    Dim lastRow As Integer
    Dim shTyp As Worksheet
    Set shTyp = ThisWorkbook.Sheets("Typschl")
    
    With shTyp
        lastRow = .UsedRange.Rows.count
        
        '' if it's not visible, then display it
        If .Visible = False And ThisWorkbook.Sheets("Home").Typschl.Value = True Then
            .Visible = True
            .Activate
            .Cells(lastRow + 1, 1).Select
            
        Else
            '' else hide it and sort on Derivat name and Typschlüssel number
            .Visible = False
            lastRow = .UsedRange.Rows.count
            .Range("A1:B" & lastRow).RemoveDuplicates Columns:=Array(1, 2) ' not sure if this doesn't include the whole table
            With .Sort
                .Header = xlYes  '' this line means this worksheet has a header row which is not included in the sort
                .SortFields.Clear
                .SortFields.add Key:=shTyp.Columns(2), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                .SortFields.add Key:=shTyp.Columns(1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                .SetRange shTyp.UsedRange
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            ThisWorkbook.Sheets("Home").Typschl.Value = False
            
        End If
    End With
End Sub

Sub checkTypschl(quelleName As String) 'As String

    Dim shQu As Worksheet, shTyp As Worksheet, shKopf As Worksheet
    Dim dataQu() As Variant, dataTyp() As Variant, dataKopf() As Variant
    Dim i As Long, j As Integer, komCol As Integer, keCol As Integer, fteCol As Integer, answer As Integer
    Dim typ As String, missingTyp As String, der As String, SOP As String, mrktsgmnt As String, gueltigkeitdatum As String, arr() As String
    Dim inTyp As Boolean
    
    '' initialize variable and arrays
    Set shQu = Workbooks(quelleName).Sheets("Strukturbericht")
    Set shKopf = Workbooks(quelleName).Sheets("Kopf mit Parameter")
    Set shTyp = ThisWorkbook.Sheets("Typschl")
    
    dataQu() = shQu.UsedRange
    dataTyp() = shTyp.UsedRange
    dataKopf() = shKopf.UsedRange
    ' The colon (:) is a statement delimiter. It would be equivalent to a new line in VBA,
    ' or a semicolon in C (just to quote a random example). It allows you to write several
    ' instructions on a single line rather than going to a new line each time.
    der = vbNullString: SOP = vbNullString: mrktsgmnt = vbNullString
    
    '' get the value of the Typschlüssel of the SAP export from the cell B16 in the array of Kopf mit Parameter
    typ = dataKopf(16, 2)
    '' cut after the dot
    typ = Right(typ, Len(typ) - InStr(typ, "."))
    
    '' get the value of the Gültigkeitdatum in cell B35, from the array of Kopf mit Parameter
    gueltigkeitdatum = dataKopf(35, 2)
'    If gueltigkeitdatum = vbNullString Then gueltigkeitdatum = "not found"
    gueltigkeitdatum = Replace(gueltigkeitdatum, "'", vbNullString)
    
    If gueltigkeitdatum = vbNullString Then
        gueltigkeitdatum = "not found"
    End If
    
    '' look up Typschlüssel from the SAP export in the Typschlüssel liste to get the match Derivat name
    inTyp = False
    For i = 1 To UBound(dataTyp, 1) '' what if der = x ?
        If dataTyp(i, 1) = typ Then
            If dataTyp(i, 2) = "x" Then
                shTyp.Rows(i).EntireRow.Delete
            Else
                inTyp = True
                shTyp.Range("G" & i) = gueltigkeitdatum
                'Test einfügen von E/EA/WA
                'shTyp.Range("L" & i) = KD
            End If
            dataTyp() = shTyp.UsedRange
            Exit For
        End If
    Next i
    
    '' if the SAP export Typschlüssel doesn't exist in the Typschlüsselliste or if it's not
    '' valid, then the user is going to have to tell the infos by himself, the user will be
    '' prompted 3 times to get the derivat name, the SOP Date and the Markt Segment
    If inTyp = False Then
        der = vbNullString: SOP = vbNullString: mrktsgmnt = vbNullString
        
        Do While der = vbNullString Or der = "Falsch"
            der = Application.InputBox("The Typschlüsselliste is not complete for the file you want to import." & vbNewLine _
                & "Please write the corresponding derivat name for Typschlüssel: " & typ & vbNewLine _
                & "Example: G01")
        Loop
        
        Do While SOP = vbNullString Or SOP = "Falsch"
            SOP = Application.InputBox("Please write the corresponding SOP for derivat: " & der & ", typschlüssel: " & typ & vbNewLine _
                & "Format: DD.MM.YYYY" & vbNewLine _
                & "Example: 28.01.2014")
            If Mid(SOP, 3, 1) = "." And Mid(SOP, 6, 1) = "." And Len(SOP) = 10 Then
                arr() = Split(SOP, ".")
                SOP = arr(1) & "/" & arr(0) & "/" & arr(2)
            Else
                MsgBox "Date not conform."
                SOP = vbNullString
            End If
        Loop
        
        Do While mrktsgmnt = vbNullString Or mrktsgmnt = "Falsch"
            mrktsgmnt = Application.InputBox("Please write the corresponding markt segment for derivat:  " & der & ", typschlüssel: " & typ & vbNewLine _
                & "Example: KKL or UKL2 or GKL")
        Loop
        
        '' the infos are written at the end of the Typschlüsselliste
        shTyp.Range("A" & UBound(dataTyp) + 1).Resize(1, 7).Value = Array(typ, der, "x", SOP, mrktsgmnt, vbNullString, gueltigkeitdatum)
        '' the updated Typschlüsselliste is stored in an array, overwriting the old data
        dataTyp() = shTyp.UsedRange
    End If
     
    '' find location of the column in the SAP export
    For i = 1 To UBound(dataQu, 1)
        For j = 1 To UBound(dataQu, 2)
            If dataQu(i, j) = "Kommunalität" Then komCol = j
            If dataQu(i, j) = "Kom. Erstverwendung" Then keCol = j
            If dataQu(i, j) = "Fzg.typ Erstverw." Then fteCol = j
        Next
    Next i
    
    '' this second loop allows to scan the Fzg.typ Bezugsteil column in the SAP export, and check if the necessary info exists in the Typschlüssel liste
    For i = 1 To UBound(dataQu, 1)
        If dataQu(i, keCol) = "NT" And (dataQu(i, komCol) = "g" Or dataQu(i, komCol) = "gSA") Then
            typ = dataQu(i, fteCol)
            inTyp = False
            
            For j = 1 To UBound(dataTyp)
                If dataTyp(j, 1) = typ Then inTyp = True: Exit For
            Next j
            
            If inTyp = False Then
                missingTyp = missingTyp & "line: " & i & " typ: " & typ & "|"
                answer = MsgBox("The Typschlüssel " & typ & " found at line " & i & " in the source file is missing from the Typschlüsselliste." & vbNewLine _
                    & "Do you want to update the Typschlüsselliste now ?" & vbNewLine _
                    & "If you don't do it now, you still can update the Typschüsselliste later", vbYesNo)
                If answer = vbYes Then
                'ask user if he want to update the typschl
                    der = vbNullString: SOP = vbNullString: mrktsgmnt = vbNullString
                    Do While der = vbNullString
                        der = Application.InputBox("The Typschlüsselliste is not complete for the file you want to import." & vbNewLine _
                            & "Please write the corresponding derivat name for Typschlüssel: " & typ & vbNewLine _
                            & "Example: G01" & vbNewLine _
                            & "You can skip this step by pressing cancel.")
                        If der = "Falsch" Then der = "x"
                    Loop
                    
                    Do While SOP = vbNullString
                        SOP = Application.InputBox("Please write the corresponding SOP for derivat: " & der & ", typschlüssel: " & typ & vbNewLine _
                            & "Format: DD.MM.YYYY" & vbNewLine _
                            & "Example: 28.01.2014" & vbNewLine _
                            & "You can skip this step by pressing cancel.")
                        If SOP = "Falsch" Then
                            SOP = "x"
                        ElseIf Mid(SOP, 3, 1) = "." And Mid(SOP, 6, 1) = "." And Len(SOP) = 10 Then
                            arr() = Split(SOP, ".")
                            SOP = arr(1) & "/" & arr(0) & "/" & arr(2)
                        Else
                            MsgBox "Date not conform."
                            SOP = vbNullString
                        End If
                    Loop
                    
                    Do While mrktsgmnt = vbNullString
                        mrktsgmnt = Application.InputBox("Please write the corresponding markt segment for derivat:  " & der & ", typschlüssel: " & typ & vbNewLine _
                            & "Example: KKL or UKL2 or GKL" & vbNewLine _
                            & "You can skip this step by pressing cancel.")
                        If mrktsgmnt = "Falsch" Then mrktsgmnt = "x"
                    Loop
                    
                    shTyp.Range("A" & UBound(dataTyp) + 1).Resize(1, 5).Value = Array(typ, der, "x", SOP, mrktsgmnt)
                    dataTyp() = shTyp.UsedRange
                Else
                    shTyp.Range("A" & UBound(dataTyp) + 1).Resize(1, 5).Value = Array(typ, "x", "x", "x", "x")
                    dataTyp() = shTyp.UsedRange
                End If
            End If
        End If
    Next i
End Sub


Sub clearTypschlValue(der As String)
    
    Dim shTyp As Worksheet
    Dim dataTyp() As Variant
    Dim i As Long

    Set shTyp = ThisWorkbook.Sheets("Typschl")
    dataTyp = shTyp.UsedRange
    
    For i = 1 To UBound(dataTyp, 1)
        If dataTyp(i, 2) = der Then
            '' gesamt value
            If dataTyp(i, 6) <> vbNullString Then
                shTyp.Cells(i, 6).ClearContents
            End If
            '' gultigkeitdatum
            If dataTyp(i, 7) <> vbNullString Then
                shTyp.Cells(i, 7).ClearContents
            End If
        End If
    Next i
   
End Sub


Function getGesamtValue(der As String) As Integer
    
    Dim pfad As String
    Dim shDer As Worksheet
    Dim gesamt As Integer
    Dim i As Long
    Dim dataDer() As Variant
    
    pfad = ThisWorkbook.Path & "\KAT_Vorlage\MEGALISTE.xlsx"
    If IsWorkBookOpen(pfad) <> True Then Workbooks.Open (pfad)
    Set shDer = Workbooks("MEGALISTE.xlsx").Sheets("Derivat")
    dataDer = shDer.UsedRange
    gesamt = 0
    
    For i = 1 To UBound(dataDer, 1)
        If dataDer(i, 1) = der Then gesamt = gesamt + 1
    Next i

    getGesamtValue = gesamt
    
End Function

Attribute VB_Name = "Allgemein_Megaliste"
Option Explicit


Sub addSAP(quelleName As String)
    '' once it's been made sure that the file is an SAP export then it's time to import it
    '' there are several steps, like checking the column position,
    '' correcting the missing information in the column of the Bezugstyp,
    '' verifying each line validity,
    '' placing it into an array with the name of the derivat in the first column,
    '' putting into the megaliste, after deleting the data for the same derivat etc...
    Dim wbMega As Workbook, wbQu As Workbook
    Dim shMega As Worksheet, shQu As Worksheet, shTyp As Worksheet
    
    Dim dataQu() As Variant, correctDataQu() As Variant, dataTyp() As Variant, dataMega() As Variant
    Dim colName() As Variant, colNumber() As Variant, colNumberMega() As Variant
    
    Dim pfad As String
    Dim der As String, typ As String, bez As String, notvalid As String
    
    Dim lastRow As Long, rw As Long
    Dim i As Integer, j As Integer, k As Integer
    
    Dim found As Boolean
    
    pfad = ThisWorkbook.Path & "\KAT_Vorlage\MEGALISTE.xlsx"
    If IsWorkBookOpen(pfad) <> True Then Workbooks.Open (pfad)
    
    Set wbMega = Workbooks("MEGALISTE.xlsx")
    Set shMega = wbMega.Sheets("Derivat")
    shMega.DisplayPageBreaks = False
    dataMega = shMega.Range(shMega.UsedRange.Rows(1).Address) '' this is the title row in shMega
    
    Set shTyp = ThisWorkbook.Sheets("Typschl")
    dataTyp = shTyp.UsedRange
    
    Set wbQu = Workbooks(quelleName)
    Set shQu = wbQu.Sheets("Strukturbericht")
    dataQu = shQu.UsedRange

    '' this array is very important: it is the link between the source data and the megaliste.
    '' if you have to add a new column to the tool i.e. a new slicer etc... please add the the new column title AT THE END of this array.
    colName() = Array("1", _
        "Objekt-Name", _
        "FB", _
        "Modulorg.", _
        "techn. Beschr.", _
        "Kom. Erstverwendung", _
        "Fzg.typ Erstverw.", _
        "Fzg.typ Bezugsteil", _
        "Beziehungswissen", _
        "BK-Cluster Text Variante", _
        "BK-Cluster Variante", _
        "Group Baukastenbezeichnung", _
        "Produkt Baukasten", _
        "Archetyp-Kateg. Variante", _
        "Prozessbaukasten", _
        "BB Teilekomm.", _
        "Kommunalität", _
        "Architekturkennzeichen", _
        "PosVar-GUID", _
        "Dimensionslosekommunalitaet", _
        "HZ1", _
        "Knoten", _
        "HZ2", _
        "HZ3", _
        "PPG")

    ReDim colNumber(LBound(colName) To UBound(colName))
    ReDim colNumberMega(LBound(colName) To UBound(colName))
    ReDim correctDataQu(LBound(dataQu) To UBound(dataQu), 1 To UBound(dataMega, 2)) '' as long as the source data, as large as the megaliste
    
    '' find column location in source file
    '' store it in the colNumber array
    For i = LBound(colName) To UBound(colName)
        found = False
        For j = 1 To UBound(dataQu)
        '' For j = 1 To 6
            For k = 1 To UBound(dataQu, 2)
                If colName(i) = dataQu(j, k) Then
                    colNumber(i) = k
                    found = True
                    Exit For
                End If
            Next k
            If found = True Then Exit For
        Next j
    '' If colNumber(i) = 0 Then MsgBox "Column " & colName(i) & "not found in the file"
    Next i
    
    '' find column location in megaliste
    '' store it in the colNumberMega array
    For i = 1 To UBound(colName)
        For j = 1 To UBound(dataMega, 2)
            If colName(i) = dataMega(1, j) Then
                colNumberMega(i) = j
                Exit For
            End If
        Next j
    '' If colNumberMega(i) = 0 Then MsgBox "Column " & colName(i) & "not found in the megaliste"
    Next i
    
    '' find derivat name :
    '' find Typschlüssel in sourcefile
    '' look up Typschlüssel in Typschlüsselliste in the tool
    typ = wbQu.Sheets("Kopf mit Parameter").Range("B16").Value '' get Typschlüssel from Excel file
    typ = Right(typ, Len(typ) - InStr(typ, ".")) ''look for it in the cell under Kommunalität
    For rw = 1 To UBound(dataTyp, 1)
        If dataTyp(rw, 1) = typ Then
            der = dataTyp(rw, 2)
            Exit For
        End If
    Next rw
    
    '' scan through each line of the source file to clean and validate data before importing into the megaliste
    For rw = 1 To UBound(dataQu, 1)
        bez = dataQu(rw, colNumber(7))
        If dataQu(rw, colNumber(16)) <> "g" And _
            dataQu(rw, colNumber(16)) <> "gSA" And _
            dataQu(rw, colNumber(16)) <> "s" And _
            dataQu(rw, colNumber(16)) <> "sSA" And _
            dataQu(rw, colNumber(16)) <> "n" And _
            dataQu(rw, colNumber(16)) <> "nSA" Then
            notvalid = notvalid & rw & "|"
        Else
            lastRow = lastRow + 1
            '' correct the data in column "Fzg.typ Bezugsteil": if a typschlüssel is found instead of a derivat, remplace it by the matching derivat
            If bez <> vbNullString Then
                bez = cleanValue(bez)
                bez = Replace(bez, " ", vbNullString)
                bez = Replace(bez, Chr(34), vbNullString)
                bez = Replace(bez, Chr(32), vbNullString)
                dataQu(rw, colNumber(7)) = bez
                For i = 1 To UBound(dataTyp, 1)
                    If dataTyp(i, 1) = bez Then
                        bez = dataTyp(i, 2)
                        dataQu(rw, colNumber(7)) = bez
                        Exit For
                    End If
                Next i
            End If
            
            '' if NT or ST is found in column "Kom. Erstverwendung" and g or gSA is found in column "Kommunalität"
            '' then get value in column "Fzg.typ Erstverw." and write matching derivat in column "Fzg.typ Bezugsteil"
            If (dataQu(rw, colNumber(5)) = "NT" Or dataQu(rw, colNumber(5)) = "ST") And (dataQu(rw, colNumber(16)) = "g" Or dataQu(rw, colNumber(16)) = "gSA") Then
                typ = dataQu(rw, colNumber(6))
                For i = 1 To UBound(dataTyp, 1)
                    If dataTyp(i, 1) = typ Then dataQu(rw, colNumber(7)) = dataTyp(i, 2): Exit For
                Next i
            End If
            
            '' writes the corrected data in a final array
            correctDataQu(lastRow, 1) = der
            correctDataQu(lastRow, 2) = dataQu(rw, colNumber(0))    '' because the name of the columns in the file and the megaliste are different
            For i = 1 To UBound(colName)
                If colNumber(i) <> 0 And colNumberMega(i) <> 0 Then '' look for column name in MEGALISTE too !
                    correctDataQu(lastRow, colNumberMega(i)) = dataQu(rw, colNumber(i))
                End If
            Next i
            
        End If
    Next rw
    
    '' clear pie chart picture from the pie folder
    Call deletePie(der)
    '' delete old data for Derivat "der" in the megaliste
    Call deleteRowMega(der, 1)
    
    '' paste new corrected data at the end of the megaliste
    '' lastrow = count valid rows see loop above
    
    shMega.Range("A" & shMega.UsedRange.Rows.count + 1).Resize(lastRow, UBound(dataMega, 2)).Value = correctDataQu '' UBound(dataMega, 2)
    
    '' write total number of parts in the Typschlüssel liste, in the 6th column
    '' this number will be used in the gesamtdarstellung
    For i = 1 To UBound(dataTyp, 1)
        If dataTyp(i, 2) = der And dataTyp(i, 7) <> vbNullString Then shTyp.Cells(i, 6) = getGesamtValue(der): Exit For '' to do: what if datatyp(i,7) = vbnullstring ?
    Next i
    
    '' If Len(notvalid) > 1 Then MsgBox Left(notvalid, Len(notvalid) - 1)
    
    shMega.DisplayPageBreaks = True

End Sub


Sub addKP(wbQu As Workbook)
' '' for test
'Sub addKP()
    '' the treatment of a Konfigurationprämissen File is slightly different from an SAP export as one file contains several Derivat.
    '' it finds the appropriate column
    '' gets the derivat names
    '' places the data in the right columns
    '' checks data validity
    '' cleans cells
    '' exports it to the megaliste
    Dim dataQu() As Variant, correctDataQu() As Variant, dataTyp() As Variant, dataMega() As Variant
    Dim colName() As Variant, colNumber() As Variant
    
    Dim destinationRow As Integer, rw As Long, lastRowMega As Long
    Dim i As Integer, j As Integer, gesamt As Integer
    
    Dim pfad As String
    Dim der As String, content As String, Kom As String, bez As String, str As String
    Dim SOP As String, mrktsgmnt As String
    Dim arr() As String, arrDate() As String, arrDer() As String
    
    Dim shQu As Worksheet, shTyp As Worksheet, shMega As Worksheet
    Dim wbMega As Workbook
    
    Dim found As Boolean
    
'    '' for test
'    Dim wbQu As Workbook, pfad_Q As String
'    pfad_Q = ThisWorkbook.Path & "\KAT_Vorlage\03_Datengrundlage\Q1 2021\210331_NCAR_KP.xlsx"
'    If IsWorkBookOpen(pfad_Q) <> True Then Workbooks.Open (pfad_Q)
'    Set wbQu = Workbooks("210331_NCAR_KP.xlsx")
'    Set shQu = wbQu.Sheets(1)
'    dataQu = shQu.UsedRange

    Set shQu = wbQu.Sheets(1)
    dataQu = shQu.UsedRange
    
    Set shTyp = ThisWorkbook.Sheets("Typschl")
    dataTyp = shTyp.UsedRange
    
    str = vbNullString
    
    '' initialize arrays
    pfad = ThisWorkbook.Path & "\KAT_Vorlage\MEGALISTE.xlsx"
    If IsWorkBookOpen(pfad) <> True Then Workbooks.Open (pfad)
    
    Set wbMega = Workbooks("MEGALISTE.xlsx")
    Set shMega = wbMega.Sheets("Derivat")
    dataMega = shMega.Range(shMega.UsedRange.Rows(1).Address) '' header row
        
    '' in old version, the parameter "FB" is not transferred to the MEGALISTE
    '' so there is no FB data, thus all data in FB show just in one column
    colName() = Array("Derivat", _
        "Komponente", _
        "Modulorg.", _
        "Fzg.typ Bezugsteil", _
        "Kommunalität", _
        "FB")
        
    ReDim colNumber(UBound(colName))
    
    '' look for column names in the megaliste, store column number in an array
    For i = LBound(colName) To UBound(colName)
        For j = LBound(dataMega, 2) To UBound(dataMega, 2)
            If dataMega(1, j) = colName(i) Then
                colNumber(i) = j
                Exit For
            End If
        Next j
    Next i
    
    '' prepare the array that will recieve the correct data
    '' its dimensions are the numbe rof derivat in the KP file times the number of lines in the KP file
    ReDim correctDataQu(1 To (UBound(dataQu) - 1) * (UBound(dataQu, 2) - 3), 1 To UBound(dataMega, 2))
    
    '' correct derivats' name
    '' the derivates are located after column 4, so we start from j=4
    For j = 4 To UBound(dataQu, 2)
        content = dataQu(1, j)
        arrDer() = Split(content, " ")
        If content <> vbNullString Then
            der = arrDer(0)
            
            If InStr(content, "NF") > 0 Then
                der = der & "NF"
            ElseIf InStr(content, "BEV") > 0 Then
                der = der & "BEV"
            ElseIf InStr(content, "PHEV") > 0 Then
                der = der & "PHEV"
            ElseIf InStr(content, "NEV") > 0 Then
                der = der & "NEV"
            Else
                der = der & "(KP)"
            End If

            dataQu(1, j) = der
            str = str & der & "|"
        End If
    Next j
    
    str = Left(str, Len(str) - 1) '' otherwise the strings ends with "|"
    
    '' for every valid cell create a new line
    For j = 4 To UBound(dataQu, 2)
        der = dataQu(1, j)
        For i = 2 To UBound(dataQu)
            content = dataQu(i, j)
            '' check column names
            If content <> vbNullString And content <> "-" Then
                content = cleanValue(content)
                content = Replace(content, " ", vbNullString)
                '' divides the cells info in 2: Kommunalität and Fahrzeug Bezugsteil
                '' clean this value for eventuel spaces or line feed
                If InStr(content, "SA") > 0 Then
                    Kom = Left(content, 3)
                    bez = Right(content, Len(content) - 3)
                Else
                    Kom = Left(content, 1)
                    bez = Right(content, Len(content) - 1)
                End If
                
                '' Es gibt Neuteile als "N" geschriben, aber nicht berechnet in FB Einzelstellung
                '' Deshalb müssen wir zuerst diese Großbuchstaben in Kleinbuchstaben wechseln
                If Kom = "G" Or Kom = "S" Or Kom = "N" Then Kom = LCase(Kom)
                
                If Kom = "g" Or Kom = "gSA" Or Kom = "s" Or Kom = "sSA" Or Kom = "n" Or Kom = "nSA" Then
                    rw = rw + 1
                    '' derivat
                    correctDataQu(rw, colNumber(0)) = der
                    '' Komponente
                    correctDataQu(rw, colNumber(1)) = dataQu(i, 2) & " - " & dataQu(i, 3)
                    '' Modulorg. : clean the value (delete non printable characters, line feed etc...)
                    content = dataQu(i, 1)
                    content = cleanValue(content)
                    correctDataQu(rw, colNumber(2)) = content
                    '' adds (KP) when there is a Bezug, to not mix up the dataset in the gesamtdarstellung
                    '' This is not wanted anymore, because there is no such ordinary data with the same derivat name
                    '' If bez <> vbNullString Then correctDataQu(rw, colNumber(3)) = bez & "(KP)"
                    correctDataQu(rw, colNumber(3)) = bez
                    '' Kommunalität
                    correctDataQu(rw, colNumber(4)) = Kom
                    
                    '' Fachbereich ~ Regel für Fachbereich ist
                    '' EV: CA
                    '' EE: CB, CD, CE
                    '' EF: Fxxx, namely, Modul begin with "F", here from FA to FG
                    '' EP: Kxxx, namely, Modul begin with "K", here from KA to KM
                    '' EA: Mxxx, namely, Modul begin with "M", here from MA to MQ
                    If InStr(dataQu(i, 1), "CA") > 0 Then
                        correctDataQu(rw, colNumber(5)) = "EV"
                    ElseIf InStr(dataQu(i, 1), "CB") > 0 Or _
                        InStr(dataQu(i, 1), "CC") > 0 Or _
                        InStr(dataQu(i, 1), "CD") > 0 Or _
                        InStr(dataQu(i, 1), "CE") > 0 Then
                        correctDataQu(rw, colNumber(5)) = "EE"
                    ElseIf InStr(dataQu(i, 1), "F") = 1 Then
                        correctDataQu(rw, colNumber(5)) = "EF"
                    ElseIf InStr(dataQu(i, 1), "K") = 1 Then
                        correctDataQu(rw, colNumber(5)) = "EP"
                    ElseIf InStr(dataQu(i, 1), "M") = 1 Then
                        correctDataQu(rw, colNumber(5)) = "EA"
                    End If
                End If
            End If
        Next i
    Next j
    
    '' str saved the names of the deriat, we cut it into an array
    '' for each derivat, we delete the existing pie and delete the lines already existing in the MEGALISTE
    arr() = Split(str, "|")
    For i = LBound(arr) To UBound(arr)
        Call deletePie(arr(i))
        Call deleteRowMega(arr(i), 1)
    Next i
    
    lastRowMega = wbMega.Sheets("Derivat").UsedRange.Rows.count
    wbMega.Sheets("Derivat").Range("A" & lastRowMega + 1).Resize(rw, UBound(dataMega, 2)) = correctDataQu
    
    '' here we fill out the Typschlüsselliste with the total number of parts per derivat
    '' if the line doesn't exist, we create it
    For i = LBound(arr) To UBound(arr)
        found = False
        der = arr(i)
        gesamt = getGesamtValue(der)
        For j = 1 To UBound(dataTyp, 1)
            If dataTyp(j, 2) = der Then shTyp.Cells(j, 6) = gesamt: shTyp.Cells(j, 7) = "x": found = True: Exit For
        Next j
        If found = False Then '' the derivat is not found in the Typschlüsselliste in the file then ask user for infos to wrtie new line for derivat in the Typschlüsselliste
            SOP = vbNullString: mrktsgmnt = vbNullString
            
            Do While SOP = vbNullString Or SOP = "Falsch"
                SOP = Application.InputBox("Please write the corresponding SOP for derivat: " & der & vbNewLine _
                    & "Format: DD.MM.YYYY" & vbNewLine _
                    & "Example: 28.01.2014")
                '' this string manipulation is there because otherwise excel exchanges date and months when pasting
                If Mid(SOP, 3, 1) = "." And Mid(SOP, 6, 1) = "." And Len(SOP) = 10 Then
                    arrDate() = Split(SOP, ".")
                    SOP = arrDate(1) & "/" & arrDate(0) & "/" & arrDate(2)
                Else
                    MsgBox "Date not conform."
                    SOP = vbNullString
                End If
            Loop
            
            Do While mrktsgmnt = vbNullString Or mrktsgmnt = "Falsch"
                mrktsgmnt = Application.InputBox("Please write the corresponding markt segment for derivat:  " & der & vbNewLine _
                    & "Example: KKL or UKL2 or GKL")
            Loop
            '' write a new line in the Typschlüsselliste
            shTyp.Range("A" & UBound(dataTyp) + 1).Resize(1, 7).Value = Array(der, der, "x", SOP, mrktsgmnt, gesamt, "x")
            '' overwrite the array with the Typschlüsselliste containning the line we just wrote
            dataTyp() = shTyp.UsedRange
        End If
    Next i
    
End Sub


Sub deleteRowMega(val As String, col As Integer)
    '' This allows to clean the megaliste from the existing data
    Dim pfad As String
    Dim wbMega As Workbook
    Dim shMega As Worksheet
    Dim dataMega() As Variant
    Dim maxi As Long
    Dim i As Long
    
    '' deletes old data from the megaliste: the keyword is the derivat cell.
    pfad = ThisWorkbook.Path & "\KAT_Vorlage\MEGALISTE.xlsx"
    If IsWorkBookOpen(pfad) <> True Then Workbooks.Open (pfad)
    
    Set wbMega = Workbooks("MEGALISTE.xlsx")
    Set shMega = wbMega.Sheets("Derivat")
    
    shMega.DisplayPageBreaks = False
    
    '' loop through the array, if the derivaname is found in the first column, delete it in the worksheet
    '' from bottom to the top
    dataMega = shMega.UsedRange
    maxi = UBound(dataMega, 1)
    For i = maxi To 1 Step -1
        If dataMega(i, col) = val Then shMega.Rows(i).EntireRow.Delete
    Next i
    
    shMega.DisplayPageBreaks = True
    
End Sub

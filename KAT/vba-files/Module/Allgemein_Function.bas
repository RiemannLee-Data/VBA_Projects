Attribute VB_Name = "Allgemein_Function"
Option Explicit


Sub deleteWs(wsName)
    
    '' delete worksheet without creating an error message
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(wsName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
End Sub


Function cleanValue(str As String) As String
    
    '' clean value in cell
    '' specialy necessary for KP import
    If Len(str) > 1000 Then str = Left(str, 1000) & "..."
    str = Application.Clean(str)
    str = Trim(str)
    str = Replace(str, Chr(13), " ")
    str = Replace(str, vbNewLine, " ")

    
    cleanValue = str
    
End Function


Sub deleteDerivat()
    
    '' deletes derivat and all small linked elements, like the pie chart saved in the folder Heatmap Kuchen Diagramm
    Dim arr As Variant
    Dim i As Integer, derCount As Integer, totDer As Integer
    Dim str As String, der As String, pfad As String
    Dim answer As Integer, answer2 As Integer
    
    If ThisWorkbook.Sheets("PIVOT").PivotTables.count = 0 Then
        MsgBox "Please import Derivat."
        Exit Sub
    End If
    
    With ThisWorkbook.Sheets("PIVOT").PivotTables("PivotTableMEGALISTE").PivotFields("Derivat")
        derCount = selectedDerivatCount
        totDer = .PivotItems.count
        
        If derCount = 1 And totDer = 1 Then
            str = .PivotItems(1).Name
        '' find number and name of unfiltered derivat
        Else
            str = vbNullString
            For i = 1 To totDer
                If .PivotItems(i).Visible = True Then
                    str = str & .PivotItems(i) & ","
                End If
            Next i
            str = Left(str, Len(str) - 1)
        End If
    End With
    
    '' send different warning message to user
    If totDer = derCount Then
        answer = MsgBox("Wollen Sie diese Derivate wirklich löschen?", vbOKCancel)
        If answer <> vbOK Then Exit Sub
    Else
        answer = MsgBox("Wollen Sie " & str & " wirklich löschen?", vbOKCancel)
        If answer <> vbOK Then Exit Sub
    End If
    
    'Derivat wird archiviert in der Historie
    answer = MsgBox("Wollen Sie das Derivat " & str & " archivieren?", vbOKCancel)
    If answer = vbOK Then
        Call archivieren(str)
    End If
    
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    '' one by one, derivat from the Typsclüsselliste, the pie chart folder, the megaliste
    arr = Split(str, ",")
    For i = 0 To UBound(arr)
        der = arr(i)
        Call deletePie(der)
        Call clearTypschlValue(der)
        Call deleteRowMega(der, 1)
    Next i

    '' reimport the megaliste in the tool
    Call createPivot
    Workbooks("MEGALISTE.xlsx").Close True

    '' display the first slicer (derivat)
    If selectedDerivatCount > 0 Then
        With ThisWorkbook
            freezeTriger = True
            .SlicerCaches.Add2(.Sheets("PIVOT").PivotTables("PivotTableMEGALISTE"), "Derivat").Slicers.add .Sheets("Home"), , "Derivat", "Derivat", 10, 180, 135, 165
            freezeTriger = False
        End With
    End If

    pfad = ThisWorkbook.Path & "\KAT_Vorlage\HISTORIE.xlsx"
    If IsWorkBookOpen(pfad) = True Then Workbooks("HISTORIE.xlsx").Activate

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    
End Sub


Function IsWorkBookOpen(fileName As String)

    '' extremely useful  and fast function
    '' check if a wb or a ppt is open
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open fileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
    
End Function


Function selectedDerivatCount() As Integer
    '' this sub is very often used. It is there to say how many derivats were selected
    '' when 1 derivat is selected, the go function will launch the Einzeldarstellung. If more, it will launch the Gesamdarstellung
    Dim i As Integer
    Dim derCount As Integer
    
    derCount = 0
    
    With ThisWorkbook.Sheets("PIVOT")
        If .PivotTables.count > 0 Then
            With .PivotTables("PivotTableMEGALISTE")
                If .PivotFields.count > 0 Then
                    If .PivotFields("Derivat").PivotItems.count = 1 Then
                        derCount = 1
                    ElseIf .PivotFields("Derivat").PivotItems.count > 1 Then
                        For i = 1 To .PivotFields("Derivat").PivotItems.count
                            If .PivotFields("Derivat").PivotItems(i).Visible = True Then
                                derCount = derCount + 1
                            End If
                        Next i
                    End If
                End If
            End With
        End If
    End With
    
    selectedDerivatCount = derCount
    
End Function

Attribute VB_Name = "Kennzahlen"
Option Explicit


Sub Kenn(quelleName As String) 'As String

    Dim shQu As Worksheet, shKD As Worksheet
    Dim i As Integer, l As Integer, z As Integer
    Dim Teilstring() As String
    Dim E As String, EA As String, WA As String
    Dim Knoten_Angabe As Boolean

    Set shQu = Workbooks(quelleName).Sheets("Strukturbericht")
    Set shKD = ThisWorkbook.Sheets("KD")

    shQu.Cells(3, 33).Value = "Dimensionslosekommunalitaet"
    shQu.Cells(3, 34).Value = "HZ1"
    shQu.Cells(3, 35).Value = "Knoten"
    shQu.Cells(3, 36).Value = "HZ2"
    shQu.Cells(3, 37).Value = "HZ3"
    
    l = shQu.Cells(Rows.count, 29).End(xlUp).row
    
    For i = 6 To l
        If shQu.Cells(i, 29) = "g" Or shQu.Cells(i, 29) = "gSA" Then
            shQu.Cells(i, 33).Value = 1
        Else
            If shQu.Cells(i, 29) = "s" Or shQu.Cells(i, 29) = "sSA" Then
                shQu.Cells(i, 33).Value = 2
            Else
                shQu.Cells(i, 33).Value = 4
            End If
        End If
    Next i
    
    Knoten_Angabe = MsgBox("Möchten Sie Knoten, welche nicht in der Referenz für Erstanläufer enthalten sind manuell nach Kommunalitätsknoten (K) und Differenzierungsknoten (D) einteilen?", vbYesNo)
    For i = 6 To l
        E = ""
        EA = ""
        WA = ""
        
        Teilstring = Split(shQu.Cells(i, 8), "/")
        shQu.Cells(i, 35) = Teilstring(1)
        
        z = ((shKD.Cells(Rows.count, 1).End(xlUp).row) + 1)
        On Error Resume Next
        E = WorksheetFunction.VLookup(shQu.Cells(i, 35), shKD.[A2:B3000], 2, False)
        If E = "" Then
            If Knoten_Angabe = vbYes Then
                E = Application.InputBox("Der Knote " & shQu.Cells(i, 35) & " ist nicht in der Referenz für Erstanläufer enthalten." & vbNewLine _
                                  & "Handelt es sich um einen Kommunalitätsknoten (K) oder einen Differenzierungsknoten (D). Beispiel: K oder D")
                shQu.Cells(i, 34) = "E" & E & "_" & shQu.Cells(i, 29)
                shKD.Cells(z, 1).Value = shQu.Cells(i, 35)
                shKD.Cells(z, 2).Value = E
            End If
        Else
            shQu.Cells(i, 34) = "E" & E & "_" & shQu.Cells(i, 29)
        End If
        
        z = ((shKD.Cells(Rows.count, 3).End(xlUp).row) + 1)
        On Error Resume Next
        EA = WorksheetFunction.VLookup(shQu.Cells(i, 35), shKD.[C2:D3000], 2, False)
        If EA = "" Then
            If Knoten_Angabe = vbYes Then
                EA = Application.InputBox("Der Knote " & shQu.Cells(i, 35) & " ist nicht in der Referenz für enge Ableitungen enthalten." & vbNewLine _
                                  & "Handelt es sich um einen Kommunalitätsknoten (K) oder einen Differenzierungsknoten (D). Beispiel: K oder D")
                shQu.Cells(i, 36) = "EA" & EA & "_" & shQu.Cells(i, 29)
                shKD.Cells(z, 3).Value = shQu.Cells(i, 35)
                shKD.Cells(z, 4).Value = EA
            End If
        Else
            shQu.Cells(i, 36) = "EA" & EA & "_" & shQu.Cells(i, 29)
        End If
        
        z = ((shKD.Cells(Rows.count, 5).End(xlUp).row) + 1)
        On Error Resume Next
        WA = WorksheetFunction.VLookup(shQu.Cells(i, 35), shKD.[E2:F3000], 2, False)
        If WA = "" Then
            If Knoten_Angabe = vbYes Then
                WA = Application.InputBox("Der Knote " & shQu.Cells(i, 35) & " ist nicht in der Referenz für weite Ableitungen enthalten." & vbNewLine _
                                  & "Handelt es sich um einen Kommunalitätsknoten (K) oder einen Differenzierungsknoten (D). Beispiel: K oder D")
                shQu.Cells(i, 37) = "WA" & WA & "_" & shQu.Cells(i, 29)
                shKD.Cells(z, 5).Value = shQu.Cells(i, 35)
                shKD.Cells(z, 6).Value = WA
            End If
        Else
            shQu.Cells(i, 37) = "WA" & WA & "_" & shQu.Cells(i, 29)
        End If
        
    Next i
    
'    If KD = "E" Then
'    For i = 6 To l
'        Teilstring = Split(shQu.Cells(i, 8), "/")
'        shQu.Cells(i, 35) = Teilstring(1)
'        shQu.Cells(i, 34) = WorksheetFunction.VLookup(shQu.Cells(i, 35), shKD.[A2:B1010], 2, False)
'        shQu.Cells(i, 34) = shQu.Cells(i, 34) & "_" & shQu.Cells(i, 29)
'    Next i
'    ElseIf KD = "EA" Then
'    For i = 6 To l
'        Teilstring = Split(shQu.Cells(i, 8), "/")
'        shQu.Cells(i, 35) = Teilstring(1)
'        shQu.Cells(i, 34) = WorksheetFunction.VLookup(shQu.Cells(i, 35), shKD.[C2:D1010], 2, False)
'        shQu.Cells(i, 34) = shQu.Cells(i, 34) & "_" & shQu.Cells(i, 29)
'    Next i
'    Else
'    For i = 6 To l
'        Teilstring = Split(shQu.Cells(i, 8), "/")
'        shQu.Cells(i, 35) = Teilstring(1)
'        shQu.Cells(i, 34) = WorksheetFunction.VLookup(shQu.Cells(i, 35), shKD.[E2:F1010], 2, False)
'        shQu.Cells(i, 34) = shQu.Cells(i, 34) & "_" & shQu.Cells(i, 29)
'    Next i
'    End If
    
End Sub

Attribute VB_Name = "Regel_Datei_Erzeugen"
'Regel_Datei_Erzeugen Modul

Public erzeugnis_format As Integer
Global Const erweiterung = ".vpx"

Sub ErzeugnisWaehlen()

' Beschreibung:
'   Zeigt Wahlform für VPX Datei Format

    Dim ufe As New UserFormErzeugnis
    
    ufe.Show
    
    If erzeugnis_format = 1 Then
        Call ErzeugnisKomplett
    ElseIf erzeugnis_format = 2 Then
        Call ErzeugnisProFB
    End If

End Sub

Sub ErzeugnisKomplett()

' Beschreibung:
'   Schafft und füllt die VPX Datei aus wenn die Option "einzige Datei" gewählt ist

    Dim makro As Worksheet
    Dim pivot As Worksheet
    Dim log As Worksheet
    
    Dim derivat As New klsDerivat
    
    Dim ordner_pfad As String
    Dim datei_pfad As String
    Dim Datum As String
    Dim Uhrzeit As String
    
    Dim i As Integer
    Dim g As Integer
    Dim s As Integer
    Dim n As Integer
    Dim o As Integer
    
    Dim grenzeAusDemMotor As Long
    Dim grenzeLeer As Long
    Dim Zeilenanzhal As Long
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    'Arbeitstabelle
    Set makro = ThisWorkbook.Worksheets("MAKRO")
    Set pivot = ThisWorkbook.Worksheets("PIVOT")
    Set log = ThisWorkbook.Worksheets("LOG")
    
    If pivot.FilterMode Then pivot.ShowAllData
    If makro.FilterMode Then makro.ShowAllData
    If log.FilterMode Then log.ShowAllData
    
    Datum = Format(Date, "yymmdd")
    Uhrzeit = Replace(Time, ":", "")
    
    Let Zeilenanzahl = pivot.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
'    On Error GoTo ErrorHandler
    
    Let grenzeAusDemMotor = pivot.Cells.Find(What:="Aus dem Motor", LookAt:=xlWhole).Row
    Let grenzeLeer = pivot.Cells.Find(What:="Leer <> EA", LookAt:=xlWhole).Row
    
    'Derivat Daten von LOG Tabelle lesen
    derivat.Lesen log
    
    'Speicherung Datei wählen
    Let ordner_pfad = OrdnerWaehlen() & "\" & Datum & "_" & Uhrzeit & "_Visualisierung_" & derivat.Name
    
'    ' For debugging
'    Let ordner_pfad = "C:\Users\q520739\Downloads" & "\" & Datum & "_" & Uhrzeit & "_Visualisierung_" & derivat.Name
    
    
    'Hauptordner erstellen
    If Dir(ordner_pfad, vbDirectory) <> "" Then
       MsgBox "Der Ordner " & ordner_pfad & " ist bereits angelegt. Löschen Sie diesen bitte!"
       Exit Sub
    Else
        MkDir ordner_pfad 'Ordner schaffen
    End If
    
    'Erstellen der XML-Datei für EA
    datei_pfad = ordner_pfad & "\Visu_" & derivat.Name & "_" & derivat.Gueltigkeitsdatum & "_" & derivat.Typschluessel & erweiterung
    Open datei_pfad For Output As #1
 
    'XML-Header
    Print #1, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #1, "<VisualReport xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""VisualReportSchema"" version=""1.0"" author=""Teamcenter Visualization 11.5.0"" date=""" & Datum & """ Time=""" & Uhrzeit & """ >"
        Print #1, "<ReportProp name=""ModOrg_Filter"" actionType=""changeAppearance"" targetParts=""visible""/>"
        'Regeln
        g = 0
        n = 0
        s = 0
        
        For i = 3 To grenzeLeer - 1
            If pivot.Cells(i, 7) = "g" Then
                g = g + 1
                Print #1, "<Rule name=""GT" & g & """>"
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    
                    '###### Regel für MC07/MD01/MD02/FE01: keine KoGr ######
                    If pivot.Cells(i, 8) = "MC07" Or pivot.Cells(i, 8) = "MD01" Or pivot.Cells(i, 8) = "MD02" Or pivot.Cells(i, 8) = "FE01" Then
                        Call PrintKeinKoGr(1, pivot, i)
                        
                    '###### normale Regel ######
                    Else
                        Call PrintNormalRegel(1, pivot, i)
                    End If
                    '###########################
                    Call PrintAction_G(1)
                Print #1, "</Rule>"
                
            ElseIf pivot.Cells(i, 7) = "n" Then
                n = n + 1
                Print #1, "<Rule name=""NT" & n & """>"
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    '###### Regel für MC07/MD01/MD02/FE01: keine KoGr ######
                    If pivot.Cells(i, 8) = "MC07" Or pivot.Cells(i, 8) = "MD01" Or pivot.Cells(i, 8) = "MD02" Or pivot.Cells(i, 8) = "FE01" Then
                        Call PrintKeinKoGr(1, pivot, i)
                    '###### normale Regel ######
                    Else
                        Call PrintNormalRegel(1, pivot, i)
                    End If
                    '###########################
                    Call PrintAction_N(1)
                Print #1, "</Rule>"
                
            ElseIf pivot.Cells(i, 7) = "s" Then
                s = s + 1
                Print #1, "<Rule name=""ST" & s & """>"
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    
                    '###### Regel für MC07/MD01/MD02/FE01: keine KoGr ######
                    If pivot.Cells(i, 8) = "MC07" Or pivot.Cells(i, 8) = "MD01" Or pivot.Cells(i, 8) = "MD02" Or pivot.Cells(i, 8) = "FE01" Then
                        Call PrintKeinKoGr(1, pivot, i)
                    '###### normale Regel ######
                    Else
                        Call PrintNormalRegel(1, pivot, i)
                    End If
                    '###########################
                    Call PrintAction_S(1)
                Print #1, "</Rule>"
            End If
        Next i
        
'###########################################################################################################################################################################################################################################################################
        '(Leer)-Komponenten aus dem Motor
        If grenzeAusDemMotor <> 0 Then
            If pivot.Cells(grenzeAusDemMotor, 7) = "g" Then
                o = 1
                g = g + 1
                
                Print #1, "<Rule name=""GT" & g & """>"
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    
                    Print #1, "<Condition operator= ""or"">"
                        Do While o <= 15
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MA" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 4
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MB" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 8
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MC" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 7
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MD" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                    Print #1, "</Condition>"
                    
                    Call PrintAction_G(1)
                Print #1, "</Rule>"
                    
            ElseIf pivot.Cells(grenzeAusDemMotor, 7) = "n" Then
                o = 1
                n = n + 1
                
                Print #1, "<Rule name=""NT" & n & """>"
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    
                    Print #1, "<Condition operator= ""or"">"
                        Do While o <= 15
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MA" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 4
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MB" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 8
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MC" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 7
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MD" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                    Print #1, "</Condition>"
                    
                    Call PrintAction_N(1)
                Print #1, "</Rule>"
                    
            ElseIf pivot.Cells(grenzeAusDemMotor, 7) = "s" Then
                o = 1
                s = s + 1
                
                Print #1, "<Rule name=""ST" & s & """>"
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    
                    Print #1, "<Condition operator= ""or"">"
                    
                        Do While o <= 15
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MA" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 4
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MB" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 8
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MC" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 7
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MD" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                    Print #1, "</Condition>"
                    
                    Call PrintAction_S(1)
                Print #1, "</Rule>"
            End If
        End If
        
'###########################################################################################################################################################################################################################################################################

        'Regel für "Keine Aussage"
        Call PrintKeineAussage(1)
        
    Print #1, "</VisualReport>"

    Close #1
'**************************************************************************************************************************************************************************************************************************************************************************************************************************
    MsgBox ("Alle Dateien wurden erstellt und unter " & ordner_pfad & " gespeichert!" & vbNewLine & "� Teamcenter Visualization 11.5.0")

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    
'ErrorHandler:
''Wenn grenzeLeer nicht gefunden wird, nimmt sie dem Wert zeilenanzahl
'    grenzeLeer = Zeilenanzahl
'    grenzeAusDemMotor = 0
'    Resume Next
End Sub

Sub ErzeugnisProFB()
'
'Beschreibung: Schafft und füllt die VPX Datei aus wenn die Option "ein Datei pro FB" gewählt ist
'
    Dim makro As Worksheet
    Dim pivot As Worksheet
    Dim log As Worksheet
    
    Dim derivat As New klsDerivat
    
    Dim ordner_pfad As String
    Dim datei_pfad As String
    Dim Datum As String
    Dim Uhrzeit As String
    
    Dim i As Integer
    Dim g As Integer
    Dim s As Integer
    Dim n As Integer
    Dim o As Integer
    
    Dim grenzeEA As Long
    Dim grenzeEE As Long
    Dim grenzeEF As Long
    Dim grenzeEP As Long
    Dim grenzeEV As Long
    Dim grenzeLeer As Long
    Dim grenzeAusDemMotor As Long
    Dim Zeilenanzahl As Long
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    'Arbeitstabelle
    Set makro = ThisWorkbook.Worksheets("MAKRO")
    Set pivot = ThisWorkbook.Worksheets("PIVOT")
    Set log = ThisWorkbook.Worksheets("LOG")
    If pivot.FilterMode Then pivot.ShowAllData
    If makro.FilterMode Then makro.ShowAllData
    If log.FilterMode Then log.ShowAllData
    
    Datum = Format(Date, "yymmdd")
    Uhrzeit = Replace(Time, ":", "")
    Let Zeilenanzahl = pivot.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    
    'Derivat
    derivat.Lesen log
    
    'Speicherung Datei wählen
    Let ordner_pfad = OrdnerWaehlen() & "\" & Datum & "_" & Uhrzeit & "_Visualisierung_" & derivat.Name
    
    'Hauptordner erstellen
    If Dir(ordner_pfad, vbDirectory) <> "" Then
       MsgBox "Der Ordner " & ordner_pfad & " ist bereits angelegt. Löschen Sie diesen bitte!"
       Exit Sub
    Else
        MkDir ordner_pfad 'Ordner schaffen
    End If

    
    grenzeEA = pivot.Cells.Find(What:="EA", LookAt:=xlWhole).Row
    grenzeEE = pivot.Cells.Find(What:="EE", LookAt:=xlWhole).Row
    grenzeEF = pivot.Cells.Find(What:="EF", LookAt:=xlWhole).Row
    grenzeEP = pivot.Cells.Find(What:="EP", LookAt:=xlWhole).Row
    grenzeEV = pivot.Cells.Find(What:="EV", LookAt:=xlWhole).Row
'    On Error GoTo ErrorHandler
    grenzeLeer = pivot.Cells.Find(What:="Leer <> EA", LookAt:=xlWhole).Row
    grenzeAusDemMotor = pivot.Cells.Find(What:="Aus dem Motor", LookAt:=xlWhole).Row
    
    'Erstellen der XML-Datei für EA
    datei_pfad = ordner_pfad & "\Visu_" & derivat.Name & "_" & derivat.Gueltigkeitsdatum & "_" & derivat.Typschluessel & "_" & "EA" & erweiterung
    Open datei_pfad For Output As #1

 
    'XML-Header für EA
    Print #1, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #1, "<VisualReport xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""VisualReportSchema"" version=""1.0"" author=""Teamcenter Visualization 11.5.0"" date=""" & Datum & """ Time = """ & Uhrzeit & """ >"
        Print #1, "<ReportProp name=""ModOrg_Filter"" actionType=""changeAppearance"" targetParts=""visible""/>"


        'Regeln für EA
        g = 0
        n = 0
        s = 0
        
        For i = grenzeEA + 1 To grenzeEE - 1
            If pivot.Cells(i, 7) = "g" Then
                g = g + 1
                Print #1, "<Rule name=""GT" & g & """>"
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    
                    '###### Regel für MC07/MD01/MD02: keine KoGr ######
                    If pivot.Cells(i, 8) = "MC07" Or pivot.Cells(i, 8) = "MD01" Or pivot.Cells(i, 8) = "MD02" Then
                        Call PrintKeinKoGr(1, pivot, i)
                    '###### normale Regel ######
                    Else
                        Call PrintNormalRegel(1, pivot, i)
                    End If
                    '###########################
                    Call PrintAction_G(1)
                Print #1, "</Rule>"
                
            ElseIf pivot.Cells(i, 7) = "n" Then
                n = n + 1
                Print #1, "<Rule name=""NT" & n & """>"
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    '###### Regel für MC07/MD01/MD02: keine KoGr ######
                    If pivot.Cells(i, 8) = "MC07" Or pivot.Cells(i, 8) = "MD01" Or pivot.Cells(i, 8) = "MD02" Then
                        Call PrintKeinKoGr(1, pivot, i)
                    '###### normale Regel ######
                    Else
                        Call PrintNormalRegel(1, pivot, i)
                    End If
                    '###########################
                    Call PrintAction_N(1)
                Print #1, "</Rule>"
                
            ElseIf pivot.Cells(i, 7) = "s" Then
                s = s + 1
                Print #1, "<Rule name=""ST" & s & """>"
                    Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    '###### Regel für MC07/MD01/MD02: keine KoGr ######
                    If pivot.Cells(i, 8) = "MC07" Or pivot.Cells(i, 8) = "MD01" Or pivot.Cells(i, 8) = "MD02" Then
                        Call PrintKeinKoGr(1, pivot, i)
                    '###### normale Regel ######
                    Else
                        Call PrintNormalRegel(1, pivot, i)
                    End If
                    '###########################
                    Call PrintAction_S(1)
                Print #1, "</Rule>"
            End If
        Next i
        
'###########################################################################################################################################################################################################################################################################
        '(Leer)-Komponenten aus dem Motor
        If grenzeAusDemMotor <> 0 Then
            If pivot.Cells(grenzeAusDemMotor, 7) = "g" Then
                o = 1
                g = g + 1
                Print #1, "<Rule name=""GT" & g & """>"
                        Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                        Print #1, "<Condition operator= ""or"">"
                        
                        Do While o <= 15
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MA" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 4
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MB" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 8
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MC" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 7
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MD" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        Print #1, "</Condition>"
                        Call PrintAction_G(1)
                    Print #1, "</Rule>"
                    
            ElseIf pivot.Cells(grenzeAusDemMotor, 7) = "n" Then
                o = 1
                n = n + 1
                Print #1, "<Rule name=""NT" & n & """>"
                        Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                        Print #1, "<Condition operator= ""or"">"
                        
                        Do While o <= 15
                        o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MA" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 4
                        o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MB" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 8
                        o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MC" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 7
                        o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MD" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        Print #1, "</Condition>"
                        Call PrintAction_N(1)
                    Print #1, "</Rule>"
                    
            ElseIf pivot.Cells(grenzeAusDemMotor, 7) = "s" Then
                o = 1
                s = s + 1
                    Print #1, "<Rule name=""ST" & s & """>"
                        Print #1, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                        Print #1, "<Condition operator= ""or"">"
                        
                        Do While o <= 15
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MA" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 4
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MB" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 8
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MC" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                        
                        o = 1
                        Do While o <= 7
                            o_print = Format(CStr(o), "00")
                            Print #1, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & "MD" & o_print & """ type= ""attribute"">"
                                Print #1, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
                            Print #1, "</Condition>"
                            o = o + 1
                        Loop
                            
                        Print #1, "</Condition>"
                        Call PrintAction_S(1)
                    Print #1, "</Rule>"
            End If
        End If
'###########################################################################################################################################################################################################################################################################


        'Regel für "Keine Aussage"
        Call PrintKeineAussage(1)
        
    Print #1, "</VisualReport>"

    Close #1
'**************************************************************************************************************************************************************************************************************************************************************************************************************************
    
    'Erstellen der XML-Datei für EE
    datei_pfad = ordner_pfad & "\Visu_" & derivat.Name & "_" & derivat.Gueltigkeitsdatum & "_" & derivat.Typschluessel & "_" & "EE" & erweiterung
    Open datei_pfad For Output As #2

 
    'XML-Header für EE
    Print #2, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #2, "<VisualReport xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""VisualReportSchema"" version=""1.0"" author=""Teamcenter Visualization 11.5.0"" date=""" & Datum & """ Time = """ & Uhrzeit & """ >"
        Print #2, "<ReportProp name=""ModOrg_Filter"" actionType=""changeAppearance"" targetParts=""visible""/>"


    'Regeln für EE
        g = 0
        n = 0
        s = 0
        
        For i = grenzeEE + 1 To grenzeEF - 1
            If pivot.Cells(i, 7) = "g" Then
                g = g + 1
                Print #2, "<Rule name=""GT" & g & """>"
                    Print #2, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Call PrintNormalRegel(2, pivot, i)
                    Call PrintAction_G(2)
                Print #2, "</Rule>"
                
            ElseIf pivot.Cells(i, 7) = "n" Then
                n = n + 1
                Print #2, "<Rule name=""NT" & n & """>"
                    Print #2, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Call PrintNormalRegel(2, pivot, i)
                    Call PrintAction_N(2)
                Print #2, "</Rule>"
                
            ElseIf pivot.Cells(i, 7) = "s" Then
                s = s + 1
                Print #2, "<Rule name=""ST" & s & """>"
                    Print #2, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Call PrintNormalRegel(2, pivot, i)
                    Call PrintAction_S(2)
                Print #2, "</Rule>"
            End If
        Next i
        
        'Regel für "Keine Aussage"
        Call PrintKeineAussage(2)
        
    Print #2, "</VisualReport>"

    Close #2
'*****************************************************************************************************************************************************************************************************************************************************************************************************
    
    'Erstellen der XML-Datei für EF
    datei_pfad = ordner_pfad & "\Visu_" & derivat.Name & "_" & derivat.Gueltigkeitsdatum & "_" & derivat.Typschluessel & "_" & "EF" & erweiterung
    Open datei_pfad For Output As #3

 
    'XML-Header für EF
    Print #3, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #3, "<VisualReport xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""VisualReportSchema"" version=""1.0"" author=""Teamcenter Visualization 11.5.0"" date=""" & Datum & """ Time = """ & Uhrzeit & """ >"
        Print #3, "<ReportProp name=""ModOrg_Filter"" actionType=""changeAppearance"" targetParts=""visible""/>"


    'Regeln für EF
        g = 0
        n = 0
        s = 0
        
        For i = grenzeEF + 1 To grenzeEP - 1
            If pivot.Cells(i, 7) = "g" Then
                g = g + 1
                Print #3, "<Rule name=""GT" & g & """>"
                    Print #3, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    '###### Regel für FE01: keine KoGr ######
                    If pivot.Cells(i, 8) = "FE01" Then
                        Call PrintKeinKoGr(3, pivot, i)
                    '###### normale Regel ######
                    Else
                        Call PrintNormalRegel(3, pivot, i)
                    End If
                    '###########################
                    Call PrintAction_G(3)
                Print #3, "</Rule>"
                
            ElseIf pivot.Cells(i, 7) = "n" Then
                n = n + 1
                Print #3, "<Rule name=""NT" & n & """>"
                    Print #3, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    '###### Regel für FE01: keine KoGr ######
                    If pivot.Cells(i, 8) = "FE01" Then
                        Call PrintKeinKoGr(3, pivot, i)
                    '###### normale Regel ######
                    Else
                        Call PrintNormalRegel(3, pivot, i)
                    End If
                    '###########################
                    Call PrintAction_N(3)
                Print #3, "</Rule>"
                
            ElseIf pivot.Cells(i, 7) = "s" Then
                s = s + 1
                Print #3, "<Rule name=""ST" & s & """>"
                    Print #3, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    '###### Regel für FE01: keine KoGr ######
                    If pivot.Cells(i, 8) = "FE01" Then
                        Call PrintKeinKoGr(3, pivot, i)
                    '###### normale Regel ######
                    Else
                        Call PrintNormalRegel(3, pivot, i)
                    End If
                    '###########################
                    Call PrintAction_S(3)
                Print #3, "</Rule>"
            End If
        Next i
        
        'Regel für "Keine Aussage"
        Call PrintKeineAussage(3)
        
    Print #3, "</VisualReport>"

    Close #3
'*****************************************************************************************************************************************************************************************************************************************************************************************************
    
   'Erstellen der XML-Datei für EP
    datei_pfad = ordner_pfad & "\Visu_" & derivat.Name & "_" & derivat.Gueltigkeitsdatum & "_" & derivat.Typschluessel & "_" & "EP" & erweiterung
    Open datei_pfad For Output As #4

 
    'XML-Header für EP
    Print #4, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #4, "<VisualReport xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""VisualReportSchema"" version=""1.0"" author=""Teamcenter Visualization 11.5.0"" date=""" & Datum & """ Time = """ & Uhrzeit & """ >"
        Print #4, "<ReportProp name=""ModOrg_Filter"" actionType=""changeAppearance"" targetParts=""visible""/>"


    'Regeln für EP
        g = 0
        n = 0
        s = 0
        
        For i = grenzeEP + 1 To grenzeEV - 1
            If pivot.Cells(i, 7) = "g" Then
                g = g + 1
                Print #4, "<Rule name=""GT" & g & """>"
                    Print #4, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Call PrintTexture(4)
                    Call PrintAction_G(4)
                Print #4, "</Rule>"
                
            ElseIf pivot.Cells(i, 7) = "n" Then
                n = n + 1
                Print #4, "<Rule name=""NT" & n & """>"
                    Print #4, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Call PrintTexture(4)
                    Call PrintAction_N(4)
                Print #4, "</Rule>"
                
            ElseIf pivot.Cells(i, 7) = "s" Then
                s = s + 1
                Print #4, "<Rule name=""ST" & s & """>"
                    Print #4, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Call PrintTexture(4)
                    Call PrintAction_S(4)
                Print #4, "</Rule>"
            End If
        Next i
        
        'Regel für "Keine Aussage"
        Call PrintKeineAussage(4)
        
    Print #4, "</VisualReport>"

    Close #4
    
'*****************************************************************************************************************************************************************************************************************************************************************************************************
    
   'Erstellen der XML-Datei für EV
    datei_pfad = ordner_pfad & "\Visu_" & derivat.Name & "_" & derivat.Gueltigkeitsdatum & "_" & derivat.Typschluessel & "_" & "EV" & erweiterung
    Open datei_pfad For Output As #6

 
    'XML-Header für EV
    Print #6, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #6, "<VisualReport xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""VisualReportSchema"" version=""1.0"" author=""Teamcenter Visualization 11.5.0"" date=""" & Datum & """ Time = """ & Uhrzeit & """ >"
        Print #6, "<ReportProp name=""ModOrg_Filter"" actionType=""changeAppearance"" targetParts=""visible""/>"


    'Regeln für EV
        g = 0
        n = 0
        s = 0
        
        For i = grenzeEV + 1 To grenzeLeer - 1
            If pivot.Cells(i, 7) = "g" Then
                g = g + 1
                Print #6, "<Rule name=""GT" & g & """>"
                    Print #6, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Call PrintTexture(6)
                    Call PrintAction_G(6)
                Print #6, "</Rule>"
                
            ElseIf pivot.Cells(i, 7) = "n" Then
                n = n + 1
                Print #6, "<Rule name=""NT" & n & """>"
                    Print #6, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Call PrintTexture(6)
                    Call PrintAction_N(6)
                Print #6, "</Rule>"
                
            ElseIf pivot.Cells(i, 7) = "s" Then
                s = s + 1
                Print #6, "<Rule name=""ST" & s & """>"
                    Print #6, "<ApplicationHint application=""TcVis"" version=""11.1""></ApplicationHint>"
                    Call PrintTexture(6)
                    Call PrintAction_S(6)
                Print #6, "</Rule>"
            End If
        Next i
        
        'Regel für "Keine Aussage"
        Call PrintKeineAussage(6)
        
    Print #6, "</VisualReport>"

    Close #6
    
'**************************************************************************************************************************************************************************************************************************************************************************************************************
    
    MsgBox ("Alle Dateien wurden erstellt und unter " & ordner_pfad & " gespeichert!" & vbNewLine & "� Teamcenter Visualization 11.5.0")

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With

'ErrorHandler:
''Wenn "Leer <> EA" nicht gefunden wird, nimmt grenzeLeer dem Wert zeilenanzahl
'    grenzeLeer = Zeilenanzahl
'    grenzeAusDemMotor = 0
'    Resume Next


End Sub




Function OrdnerWaehlen() As String
    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    fd.AllowMultiSelect = False 'Nur ein Ordner wählen
    fd.Title = "wählen Sie das Ordner wo die Regel Datei speichern"
    
    'Wenn der Benutzer auf "OK" Schaltfl�che dr�ckt
    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
        OrdnerWaehlen = vrtSelectedItem
        
    Next
    'Wenn der Benutzer auf "Abbrechen" Schaltfl�che dr�ckt
    Else
        OrdnerWaehlen = ThisWorkbook.Path
    End If
    
    Set fd = Nothing
End Function




Sub PrintKeinKoGr(FileName As Integer, PivotSheet As Worksheet, IterNum As Integer)
    Print #FileName, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & PivotSheet.Cells(IterNum, 8) & """ type= ""attribute"">"
        Print #FileName, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
    Print #FileName, "</Condition>"
End Sub




Sub PrintNormalRegel(FileName As Integer, PivotSheet As Worksheet, IterNum As Integer)
    
    Print #FileName, "<Condition operator= ""and"">"
        Print #FileName, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & PivotSheet.Cells(IterNum, 8) & """ type= ""attribute"">"
            Print #FileName, "<Property key=""ud_PDM_MODULE_MO"" type=""jtProperty""/>"
        Print #FileName, "</Condition>"
        
        Print #FileName, "<Condition caseSensitivity = ""false"" operator=""equalTo"" value= """ & Format(CStr(PivotSheet.Cells(IterNum, 1)), "0000") & """ type= ""attribute"">"
            Print #FileName, "<Property key=""ud_PDM_KOGR"" type=""jtProperty""/>"
        Print #FileName, "</Condition>"
    Print #FileName, "</Condition>"
End Sub




Sub PrintAction_G(FileName As Integer)
    Print #FileName, "<Action type=""matched"" displayMode=""solid wireframe"">"
    Print #FileName, "<SimpleClassifier name=""Aktion"">"
        Print #FileName, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
            Print #FileName, "<BasicMaterial diffuse=""0.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.000000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
            Call PrintTexture(FileName)
        Print #FileName, "</Material>"
    Print #FileName, "</SimpleClassifier>"
    Print #FileName, "</Action>"
End Sub




Sub PrintAction_N(FileName As Integer)
    Print #FileName, "<Action type=""matched"" displayMode=""solid wireframe"">"
    Print #FileName, "<SimpleClassifier name=""Aktion"">"
        Print #FileName, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
            Print #FileName, "<BasicMaterial diffuse=""1.000000 0.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.000000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
            Call PrintTexture(FileName)
        Print #FileName, "</Material>"
    Print #FileName, "</SimpleClassifier>"
    Print #FileName, "</Action>"
End Sub




Sub PrintAction_S(FileName As Integer)
    Print #FileName, "<Action type=""matched"" displayMode=""solid wireframe"">"
    Print #FileName, "<SimpleClassifier name=""Aktion"">"
        Print #FileName, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
            Print #FileName, "<BasicMaterial diffuse=""1.000000 1.000000 0.000000"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.465000 0.465000 0.000000"" transparency=""0.000000"" shininess=""0.300000""/>"
            Call PrintTexture(FileName)
        Print #FileName, "</Material>"
    Print #FileName, "</SimpleClassifier>"
    Print #FileName, "</Action>"
End Sub




Sub PrintKeineAussage(FileName As Integer)
    Print #FileName, "<Action type=""nonMatched"" displayMode=""solid wireframe"">"
        Print #FileName, "<SimpleClassifier>"
            Print #FileName, "<Material name="""" colorOn=""true"" texturesOn=""false"" bumpMapOn=""false"" envMapOn=""true"" type=""advanced"">"
                Print #FileName, "<BasicMaterial diffuse=""0.498039 0.498039 0.498039"" specular=""0.410000 0.410000 0.410000"" emissive=""0.000000 0.000000 0.000000"" ambient=""0.231588 0.231588 0.231588"" transparency=""0.750000"" shininess=""0.300000""/>"
                Call PrintTexture(FileName)
            Print #FileName, "</Material>"
        Print #FileName, "</SimpleClassifier>"
    Print #FileName, "</Action>"
End Sub




Sub PrintTexture(FileName As Integer)
    
    Print #FileName, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""0"">"
        Print #FileName, "<Matrix> 1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000 </Matrix>"
    Print #FileName, "</Texture>"
    
    Print #FileName, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""1"">"
        Print #FileName, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
    Print #FileName, "</Texture>"
    
    Print #FileName, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""2"">"
        Print #FileName, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
    Print #FileName, "</Texture>"
    
    Print #FileName, "<Texture textureOn=""false"" blendColor=""1.000000 1.000000 1.000000"" borderColor=""1.000000 1.000000 1.000000"" transparencyColor = ""0.000000 0.000000 0.000000"" textureStage=""3"">"
        Print #FileName, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
    Print #FileName, "</Texture>"
    
    Print #FileName, "<EnvMap layer=""3"" blendColor=""1.000000 1.000000 1.000000"" captureCameraPosition=""0.000000 0.000000 0.000000"">"
        Print #FileName, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
    Print #FileName, "</EnvMap>"
    
    Print #FileName, "<BumpMap>"
        Print #FileName, "<Matrix>1.000000 0.000000 0.000000 0.000000 1.000000 0.000000 0.000000 0.000000 1.000000</Matrix>"
    Print #FileName, "</BumpMap>"
End Sub

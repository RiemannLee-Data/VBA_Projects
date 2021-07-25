Attribute VB_Name = "Allgemein_Export_PowerPoint"
Option Explicit



Sub exportppt(EG As String)

    Dim i As Integer
    Dim answer As Integer
    '' selecting PowerPoint and pressing ., with all objects related to the Powerpoint Object Library available
    Dim ppPr As PowerPoint.Presentation
    Dim ppSl As PowerPoint.Slide
    Dim ppApp As PowerPoint.Application
    Dim ppLa As CustomLayout
    Dim ppName As String
    
    '' Import the powerpoint library
    Call importLibrary
    
    '' Objektzuweisung erfordern die Set-Anweisung
    '' New erstellt dann das eigenliche Objekt (eine neue Instanz der Klasse)
    Set ppApp = New PowerPoint.Application
    
    '' ask user where to save presentation
    answer = MsgBox("You are about to save those graphs in a powerpoint file." & vbNewLine & _
        "Would you like to add this content to an existing presentation ?" & vbNewLine & _
        "If you press no, a new presentation will be created.", _
        vbYesNoCancel + vbQuestion, "Powerpoint Presentation")
    
    '' verify existence of template
    If answer = vbNo Then
        If Dir(ThisWorkbook.Path & "\KAT_Vorlage\Vorlage_PowerPointExport.pptx") = vbNullString Then
            MsgBox "please add Vorlage Einzeldarstellung.pptx to this directory."
        Else
            ppName = openTemplate ''openTemplate refer to the function below
            If ppName = vbNullString Or ppName = "Falsch" Then Exit Sub
        End If
        
    '' find existing presentation
    ElseIf answer = vbYes Then
        ppName = Application.GetOpenFilename("Presentation Files (*.pptx),*.pptx", , "Please choose a file to open")
        If ppName = vbNullString Or ppName = "Falsch" Then Exit Sub
        
        If IsWorkBookOpen(ppName) = "Falsch" Then
            ppApp.Presentations.Open (ppName)
        End If
        ppName = Right(ppName, Len(ppName) - InStrRev(ppName, "\"))
        
        '' locate presentation
        For i = 1 To ppApp.Presentations.count
            If ppApp.Presentations(i).Name = ppName Then
                Set ppPr = ppApp.Presentations(i)
                Set ppSl = ppPr.Slides.AddSlide(ppPr.Slides.count + 1, ppPr.Slides(ppPr.Slides.count).CustomLayout)
                Exit For
            End If
        Next i
    Else
        Exit Sub
    End If
    
    '' start importing charts to presentation
    If EG = "einzel" Then
        Call updatePresentationEinzel(ppName)
    ElseIf EG = "gesamt" Then
        Call updatePresentationGesamt(ppName)
    End If
    
End Sub


'' import ppt library
'' Remember that user must add Macro to the Trust Center of the excel
'' Otherwise may show some errors like: Der programische Zugriff auf das Visual Basic-Projekt ist nicht sicher
Sub importLibrary()

    Dim VBEObj As Object
    
    Set VBEObj = Application.VBE.ActiveVBProject.References
    
    On Error Resume Next
    VBEObj.AddFromFile "MSPPT.OLB" 'das ist die Powerpoint Library
    
End Sub


Function openTemplate() As String

    Dim ppApp As PowerPoint.Application
    Dim ppPr As PowerPoint.Presentation
    Dim pfad As String, str As String, i As Integer
    
    openTemplate = vbNullString
    
    pfad = ThisWorkbook.Path
    Set ppApp = New PowerPoint.Application
    
    '' ask user for a ppt name
    str = Application.GetSaveAsFilename(Format(Date, "YYYY-MM-DD") & "_KAT_Auswertung", "Presentation Files (*.pptx),*.pptx")
    
    
    ''
    If str <> vbNullString And str <> "Falsch" Then
        If IsWorkBookOpen(pfad & "\KAT_Vorlage\Vorlage_PowerPointExport.pptx") = "Falsch" Then
            ppApp.Presentations.Open (pfad & "\KAT_Vorlage\Vorlage_PowerPointExport.pptx")
        End If
        
        '' locate presentation and save as
        For i = 1 To ppApp.Presentations.count
            If ppApp.Presentations(i).Name = "Vorlage_PowerPointExport.pptx" Then
                Set ppPr = ppApp.Presentations(i)
                openTemplate = True
                ppPr.SaveAs str
                openTemplate = Right(str, Len(str) - InStrRev(str, "\"))
                Exit For
            End If
        Next i
    End If

    
End Function


Sub updatePresentationEinzel(ppName As String)

    Dim ppApp As PowerPoint.Application
    Dim ppPr As PowerPoint.Presentation
    Dim ppSl As PowerPoint.Slide
    Dim str As String, gueltigkeitdatum As String, der As String
    Dim datestr As String
    Dim shTyp As Worksheet
    Dim dataTyp() As Variant
    Dim i As Integer
    
    Set shTyp = ThisWorkbook.Sheets("Typschl")
    dataTyp = shTyp.UsedRange
    
    '' find gültigkeit Datum in the Typschlüsselliste
    der = ThisWorkbook.Sheets("Home").ChartObjects("pieDia").Chart.ChartTitle.Caption
    gueltigkeitdatum = vbNullString
    
    For i = 1 To UBound(dataTyp)
        If dataTyp(i, 2) = der And dataTyp(i, 7) <> vbNullString Then
            gueltigkeitdatum = dataTyp(i, 7)
        End If
    Next i
    
    '' locate presentation and last slide
    Set ppApp = New PowerPoint.Application
    For i = 1 To ppApp.Presentations.count
        If ppApp.Presentations(i).Name = ppName Then
            Set ppPr = ppApp.Presentations(i)
            Set ppSl = ppPr.Slides(ppPr.Slides.count)
            Exit For
        End If
    Next i
    
    
    '' fill in text area on slide
    ppSl.Shapes.Item(1).TextFrame.TextRange.Text = der & " - KAT Auswertung"
    datestr = ppSl.Shapes.Item(2).TextFrame.TextRange.Text
    ppSl.Shapes.Item(2).TextFrame.TextRange.Text = Replace(datestr, "Datum", Date)
    str = "Gültigkeitsdatum PDM Export: " & gueltigkeitdatum & vbNewLine & ThisWorkbook.Sheets("Home").Range("A13")
    
    With ppSl.Shapes.Item(4).TextFrame.TextRange
        .Text = str
        .ParagraphFormat.Bullet.Visible = False
        .Font.size = 8
    End With
    
    '' Kopieren der Summary Tabelle
    Worksheets("Home").ListObjects("ZusTab").Range.CopyPicture
    ppSl.Shapes.Paste
    
    With ppSl.Shapes(5)
        .Left = 30
        .Top = 240
        .Height = 10
        .Width = 420
    End With
    
    '' Kopieren des Kuchendiagramms
    Worksheets("Home").ChartObjects("pieDia").Chart.CopyPicture
    ppSl.Shapes.Paste
    With ppSl.Shapes(6)
        .Left = 30
        .Top = 290
        .Height = 190
    End With
    
    '' Kopieren des Stufendiagramm Top 5
    Worksheets("Home").ChartObjects("trepDia").Chart.CopyPicture
    ppSl.Shapes.Paste
    With ppSl.Shapes(7)
        .Left = 258
        .Top = 290
        .Height = 190
    End With
    
    '' Kopieren des Stufendiagramm FB
    Worksheets("Home").ChartObjects("trepFB").Chart.CopyPicture
    ppSl.Shapes.Paste
    With ppSl.Shapes(8)
        .Left = 603
        .Top = 290
        .Height = 190
    End With

    '' save presentation
    If Not ppPr.Saved And ppPr.Path <> vbNullString Then ppPr.Save
    
End Sub


Sub updatePresentationGesamt(ppName As String)

    Dim ppApp As PowerPoint.Application
    Dim ppPr As PowerPoint.Presentation
    Dim ppSl As PowerPoint.Slide
    Dim shTyp As Worksheet
    Dim datestr As String
    Dim dataTyp() As Variant
    Dim i As Integer
    Dim Shp As Shape
    
    Set shTyp = ThisWorkbook.Sheets("Typschl")
    dataTyp = shTyp.UsedRange
    
    '' locate presentation and last slide
    Set ppApp = New PowerPoint.Application
    For i = 1 To ppApp.Presentations.count
        If ppApp.Presentations(i).Name = ppName Then
            Set ppPr = ppApp.Presentations(i)
            Set ppSl = ppPr.Slides(ppPr.Slides.count)
            Exit For
        End If
    Next i
    
    '' fill text area in last slide
    ppSl.Shapes.Item(1).TextFrame.TextRange.Text = "Gesamt Auswertung"
    datestr = ppSl.Shapes.Item(2).TextFrame.TextRange.Text
    ppSl.Shapes.Item(2).TextFrame.TextRange.Text = Replace(datestr, "Datum", Date)
    
    With ppSl.Shapes.Item(4).TextFrame.TextRange
        .Text = ThisWorkbook.Sheets("Home").Range("A41")
        .ParagraphFormat.Bullet.Visible = False
        .Font.size = 8
    End With
    
    '' Kopieren der Summary Tabelle
    Worksheets("Home").ChartObjects("HeatMap").Chart.CopyPicture
    ppSl.Shapes.Paste
    With ppSl.Shapes(5)
        .Left = 30
        .Top = 190
        .Height = 300
    End With
    
    '' Kopieren des Scoringdiagramms
    Worksheets("Home").ChartObjects("ScoringDia").Chart.CopyPicture
    ppSl.Shapes.Paste
    With ppSl.Shapes(6)
        .Left = 500
        .Top = 190
        .Height = 300
    End With
    
''''''''''''''''''''''''''''' New graphs since 24.06.2021 '''''''''''''''''''''''''''''
    '' Kopieren des Kuchendiagramms
    Worksheets("Home").ChartObjects("pieDiaGesamt").Chart.CopyPicture
    ppSl.Shapes.Paste
    With ppSl.Shapes(7)
        .Left = 400
        .Top = 15
        .Height = 170
    End With
    
    '' Kopieren des Trepdiagramms
    Worksheets("Home").ChartObjects("trepGesamt").Chart.CopyPicture
    ppSl.Shapes.Paste
    With ppSl.Shapes(8)
        .Left = 600
        .Top = 15
        .Height = 170
    End With
    
    '' save presentation
    If Not ppPr.Saved And ppPr.Path <> vbNullString Then ppPr.Save

End Sub

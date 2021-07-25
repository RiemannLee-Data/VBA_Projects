VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'' Jede Button im worksheet "Home" entspricht eine Funktion here in Tabelle 3


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    '' besically first check to see if we find ourselves in the Einzeldarstellung config
    If selectedDerivatCount = 1 And Target <> vbNullString And Target <> 0 Then
        Call detailEinzel(Target)
    End If
End Sub

'' öffnen die KAT-Gebrauchsenleitung in andere Folder wenn es nüztlich ist
Private Sub CommandButton3_Click()

    CommandButton3.AutoSize = False
    Dim pfad As String

    pfad = ThisWorkbook.Path & "\KAT_Vorlage\04_Gebrauchsanleitung\KAT-Gebrauchsanleitung.pdf"

    ActiveWorkbook.FollowHyperlink pfad

End Sub


Private Sub SlicerChecklist_Click()
    '' refills the checklist everytime it is displayed again
    '' so that the nummber of items in the list box always corresponds to the pivotfields
    SlicerChecklist.AutoSize = False
    If selectedDerivatCount > 0 Then
        Call Checklist.ListBox1_Initialize
        Checklist.Show
    Else
        MsgBox "Please import derivat."
    End If
End Sub


Private Sub slicerReset_Click()
    slicerReset.AutoSize = False
    '' resets all slicer so that no filter is set.
    '' this is good to make sure that also hidden slicers don't alter the results
    Call resetSlicer
End Sub


Private Sub pivotRefresh_Click()

    pivotRefresh.AutoSize = False
    '' allows to relauch the graph updates without reselecting all filters
    Call go
End Sub


Private Sub Typschl_Click()

    Typschl.AutoSize = False
    '' displays the Typschlüsselliste which normally hidden for maintenance
    Call displayTypschl
End Sub


Private Sub ImportFile_Click()

    ImportFile.AutoSize = False
    '' allows user to import new derivats into the tool
    Call importnewdata
End Sub


'' this allows to attribute a sub-module for each button
Private Sub deleteDer_Click()

    deleteDer.AutoSize = False
    Call deleteDerivat
End Sub


Private Sub nurBasis_Click()

    nurBasis.AutoSize = False
    '' this button is there to display or not display the SA in the Einzeldarstellung
    '' if there is more than one derivat chosen in the slicers, message the user that it's not possible to launch the button and reset it to false
    '' else call go, which launches Einzeldarstellung and will take into account that we only want the non-SA.
    If selectedDerivatCount <> 1 Then
        If ThisWorkbook.Sheets("Home").nurBasis = True Then
            MsgBox "The current filters and the Einzeldarstellung filters do not match." & vbNewLine & "Please reselect the same filters to run this command."
            ThisWorkbook.Sheets("Home").nurBasis = False
        End If
    Else
        Call go
    End If
End Sub


Private Sub ppt_Click()

    ppt.AutoSize = False
    '' creates an excel export for the graphs in the Einzeldarstellung
    Call exportppt("einzel")
End Sub


Private Sub SavePie_Click()

    SavePie.AutoSize = False
    '' creates a impage of the displayed pie chart that will be used in the Gesamtdarstellung
    Call saveAsPicture
End Sub


Private Sub btn_Visualisierung_Click()
    
    btn_Visualisierung.AutoSize = False
    Call Datenimport
End Sub


Private Sub detailMode_Click()

    detailMode.AutoSize = False
    '' This calls a function that attributes a macro on the arrows of the Gesamtdarstellung.
    '' when there are a lot of arrows, the attributions takes up to 5s. Therefore it is lauched independantly, not as teh arrows are created.
    '' the macro produces a worksheet which the detailled view of what transits from one derivat to another.
    '' the button also allows the user to not trigger this event, because it works a bit too well and creates a table, everytime an arrow is clicked on, which is annoying.
    If selectedDerivatCount = 0 Then
        If ThisWorkbook.Sheets("Home").detailMode = True Then
            MsgBox "There are no available derivat in the tool."
            ThisWorkbook.Sheets("Home").detailMode = False
        End If
    Else
        Call initializeDetail
    End If
End Sub


Private Sub shapeButton_Click()

    shapeButton.AutoSize = False
    '' hides pie in the Gesamtdarstellung because thez can alter visibility
    Call hidePie
End Sub


Private Sub pptGesamt_Click()

    pptGesamt.AutoSize = False
    '' creates an excel export for the graphs in the Gesamtdarstellung
    Call exportppt("gesamt")
End Sub

Private Sub Visualisierung_Click()
    ''Erzeugt XML-Datei zum Einlesen in VisView
    ''widersinnig ist, dass man kann diese Funktion "CAD" nicht finden in alle Module
    ''Also gibt es keine Information wenn man Strg+I drucken
    Call CAD
End Sub




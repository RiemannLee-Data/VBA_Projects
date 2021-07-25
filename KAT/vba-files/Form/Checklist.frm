VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Checklist 
   Caption         =   "Please choose a slicer"
   ClientHeight    =   4875
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   3730
   OleObjectBlob   =   "Checklist.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Checklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Checklist.Hide
End Sub
Private Sub CommandButton2_Click()
    Call clearSlicer
End Sub

Sub ListBox1_Click()
    Call updSl(ListBox1)
End Sub

Sub ListBox1_Initialize()

    Dim i As Integer
    ListBox1.Clear
    With ThisWorkbook.Sheets("Pivot").PivotTables("PivotTableMEGALISTE")
        For i = 1 To .PivotFields.count
            '' we don't want to display the Kommunalität Slicer because is is necessary for the rest of the algorithm.
            '' we don't want to dispaly the other slicers because they can be useful for the detailliste, but not in the filtering process (plus they take a lot of calculation power)
            If .PivotFields(i) <> "Kommunalität" And _
                .PivotFields(i) <> "Objekt-Name" And _
                .PivotFields(i) <> "Dimensionslosekommunalitaet" And _
                .PivotFields(i) <> "HZ1" And _
                .PivotFields(i) <> "Beziehungswissen" And _
                .PivotFields(i) <> "Fzg.typ Erstverw." And _
                .PivotFields(i) <> "PosVar-GUID" And _
                .PivotFields(i) <> "techn. Beschr." Then
                Checklist.ListBox1.AddItem .PivotFields(i)
            End If
        Next i
    End With
End Sub

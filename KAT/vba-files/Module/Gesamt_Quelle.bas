Attribute VB_Name = "Gesamt_Quelle"
Option Explicit

'' this quelle tab is the table used to supply the information for the Gesamtdarstellung graph with arrows.
'' its devided into 4 columns: the derivat name, the markt segment, the SOP date and the y axis value
Sub quelletab()
        
    Dim shTyp As Worksheet
    Dim tbl As ListObject
    Dim lastnkt As String
    Dim pos As Integer, i As Integer, rw As Integer, increment As Integer, k As Integer, min As Integer
    Dim j As Long
    Dim dataTyp() As Variant, arr() As Variant, mkt() As Variant
    Dim inMkt As Boolean
    
    Set tbl = ThisWorkbook.Sheets("Home").ListObjects("quelleTab")
    Set shTyp = ThisWorkbook.Sheets("Typschl")
    dataTyp = shTyp.UsedRange
    
    '' clean table and resize it to the number of selected derivat (first the info is stored in array then written in the table)
    tbl.DataBodyRange.ClearContents
    rw = selectedDerivatCount
    ReDim arr(1 To rw, 1 To 4)
    
    rw = 0
    With ThisWorkbook.Sheets("PIVOT").PivotTables("PivotTableMEGALISTE")
        For i = 1 To .PivotFields("Derivat").PivotItems.count
            If .PivotFields("Derivat").PivotItems(i).Visible = True Then
                rw = rw + 1
                arr(rw, 1) = .PivotFields("Derivat").PivotItems(i).Name
                For j = 1 To UBound(dataTyp)
                    '' the SOP and Markt Segment are read from the Typschl table
                    If dataTyp(j, 2) = arr(rw, 1) And dataTyp(j, 7) <> vbNullString Then '' if gultigkeitdatum is filed
                        arr(rw, 2) = dataTyp(j, 4) 'SOP
                        arr(rw, 3) = dataTyp(j, 5) 'MarktSegment
                    End If
                Next j
            End If
        Next i
    End With
    
    tbl.Resize Range("quelleTab[#All]").Resize(rw + 1, 4)
    tbl.DataBodyRange.Value = arr
    
    '' the table is sorted by Markt Segment then by date in order to generate "wert": the y-axis value
    Call sortquelle
    Call generateWert
    
End Sub


Sub sortquelle()

    Dim tbl As ListObject
    
    Set tbl = ThisWorkbook.Sheets("Home").ListObjects("quelleTab")
    With tbl.Sort
        .SortFields.Clear
    '' the Markt Segment Sorting has to be manual,
    '' so that UKL1 comes first, GKL last so that,
    '' on the chart it will be place bottom to top
        .SortFields.add Key:=Range("quelleTab[Markt Segment]"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, _
            CustomOrder:="UKL1,UKL2,KKL,MKL,GKL", _
            DataOption:=xlSortNormal
        .SortFields.add Key:=Range("quelleTab[SOP]"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub


Sub generateWert()

    Dim i As Integer, j As Integer, increment As Integer
    Dim wert As Double
    Dim tbl As ListObject
    Dim sameSop As Boolean
    Dim mkt As String
    
    Set tbl = ThisWorkbook.Sheets("Home").ListObjects("quelleTab")
    '' this is a bit tricky, but the goal is to spread the point on the chart, because, if the point are aligned, then the arrows will overlap each other
    '' to achieve this we used a sinus function in which we input the value of the date (i.e.: 24.12.2000 = 36884)
    '' if the points have the same sop, then they would overlap on each other
    '' soif they have the same sop, we just put the 2nd above the first
    '' the points also have to be separated in group (the markt segments)
    '' to achieve this we use an increment which jumps for every new markt segment
    '' the increment start at -3 because -3 + 4 = 1 which make a nice offset of 1 for the first point, so that it's not directly on the x-axis
    '' the +1 is because sin goes from -1 to 1 and we want it to go from 0 to 2
    increment = -3
    mkt = vbNewLine
    
    For i = 1 To tbl.DataBodyRange.Rows.count
        sameSop = False
        If tbl.DataBodyRange(i, 3) <> mkt Then
            increment = increment + 4
            wert = increment
            mkt = tbl.DataBodyRange(i, 3)
        End If
        
        If i > 1 Then
            If tbl.DataBodyRange(i - 1, 2) = tbl.DataBodyRange(i, 2) Then
                tbl.DataBodyRange(i, 4) = tbl.DataBodyRange(i - 1, 4) + 0.5
                sameSop = True
            End If
        End If
        
        If sameSop = False Then tbl.DataBodyRange(i, 4) = wert
        wert = increment + 1 + Sin((tbl.DataBodyRange(i, 2)))
    Next i
    
End Sub



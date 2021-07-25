VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'' this sub is very important
'' it allows the automatic update of the charts
Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    '' freezeTriger is there to keep the pivot from updating
    '' on each manual pivot change, the freezeTriger has to be set en true, then back on false
    '' manual changes are mostly stored in the module Pivot_Switch
    '' freezetriger is a public boolean declared in Pivot_Switch
    If freezeTriger = False Then
        Call go
    End If

End Sub

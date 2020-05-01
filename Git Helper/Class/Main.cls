VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub clearMain()
    Union(Range("N11:X12"), Range("N13:X14"), Range("M18:X20"), Range("M25:X27"), Range("M32:X34"), Range("M39:X41")).ClearContents
End Sub

Public Sub clearSheet(WS As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = WS.Cells(Rows.Count, "A").End(xlUp).Row
    lastCol = WS.Cells(5, Columns.Count).End(xlToLeft).Column

    WS.Range(WS.Cells(2, 1), WS.Cells(lastRow, lastCol)).ClearContents
End Sub


Sub gotoMain()
    Main.Select
End Sub

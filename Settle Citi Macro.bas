Attribute VB_Name = "Module_CITI"
Option Explicit

Sub RunCitiAggregation()
    Dim ws As Worksheet, n As Long, s As Long
    Set ws = ThisWorkbook.Worksheets("General")
    n = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row
    ws.Range("T" & (n + 1) & ":U" & (ws.Cells(ws.Rows.Count, "U").End(xlUp).Row + 6)).ClearContents
    ws.Range("W" & (n + 1) & ":Y" & (n + 6)).ClearContents
    ws.Range("I" & (n + 1) & ":I" & (n + 6)).ClearContents
    ws.Range("T2:T" & n).Formula = "=IF(Y2=0,O2,IF(Y2<15,15,Y2))"
    ws.Range("U2:U" & n).Formula = "=(T2-W2)*IF(ISNUMBER(SEARCH(""LONG CORRIDOR"",$F2)),0.5,0.4)"
    ws.Range("W2:W" & n).Formula = "=IF(D2=270175,21,IF(Q2<>0,0,16))"
    ws.Range("X2:X" & n).Formula = "=VLOOKUP(V2,ClientListing!E:F,2,FALSE)"
    ws.Range("Y2:Y" & n).Formula = "=O2"
    s = n + 1
    ws.Range("J" & s).Formula = "=SUM(J2:J" & n & ")"
    ws.Range("O" & s).Formula = "=SUM(O2:O" & n & ")"
    ws.Range("U" & s).Formula = "=SUM(U2:U" & n & ")"
    ws.Range("I" & (s + 3)).Formula = "=U" & s
    On Error Resume Next
    ThisWorkbook.Names("MTD_HandOff").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="MTD_HandOff", RefersTo:="='" & ws.Name & "'!" & ws.Range("I" & (s + 3)).Address
End Sub

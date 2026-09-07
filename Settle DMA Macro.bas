Attribute VB_Name = "Module_DMA"
Option Explicit

Sub RunDmaAggregation()
    Dim ws As Worksheet, n As Long, s As Long
    Set ws = ThisWorkbook.Worksheets("General")
    n = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row
    ws.Range("T" & (n + 1) & ":W" & (ws.Cells(ws.Rows.Count, "T").End(xlUp).Row + 6)).ClearContents
    ws.Range("I" & (n + 1) & ":I" & (n + 6)).ClearContents
    ws.Range("T2:T" & n).Formula = "=ROUND(J2*0.005,2)"
    ws.Range("U2:U" & n).Formula = "=ROUND(O2-T2,2)"
    ws.Range("V2:V" & n).Formula = "=ROUND(U2-W2,2)"
    ws.Range("W2:W" & n).Value = 5
    s = n + 1
    ws.Range("J" & s).Formula = "=SUM(J2:J" & n & ")"
    ws.Range("O" & s).Formula = "=SUM(O2:O" & n & ")"
    ws.Range("T" & s).Formula = "=SUM(T2:T" & n & ")"
    ws.Range("U" & s).Formula = "=SUM(U2:U" & n & ")"
    ws.Range("V" & s).Formula = "=SUM(V2:V" & n & ")"
    ws.Range("W" & s).Formula = "=SUM(W2:W" & n & ")"
    ws.Range("I" & (s + 3)).Formula = "=V" & s
    On Error Resume Next
    ThisWorkbook.Names("MTD_HandOff").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="MTD_HandOff", RefersTo:="='" & ws.Name & "'!" & ws.Range("I" & (s + 3)).Address
End Sub

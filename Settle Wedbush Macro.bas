Attribute VB_Name = "Module_Wedbush"
Option Explicit

Sub RunWedbushAggregation()
    Dim ws As Worksheet, n As Long, s As Long
    Set ws = ThisWorkbook.Worksheets("General")
    n = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row
    ws.Range("T" & (n + 1) & ":Y" & (ws.Cells(ws.Rows.Count, "T").End(xlUp).Row + 6)).ClearContents
    ws.Range("I" & (n + 1) & ":I" & (n + 6)).ClearContents
    ws.Range("T2:T" & n).Formula = "=O2"
    ws.Range("U2:U" & n).Formula = "=(T2-W2)*IF(ISNUMBER(SEARCH(""LONG CORRIDOR"",$F2)),0.5,0.6)"
    ws.Range("V2:V" & n).Formula = "=(T2-W2)*IF(ISNUMBER(SEARCH(""LONG CORRIDOR"",$F2)),0.5,0.4)"
    ws.Range("W2:W" & n).Value = 5
    ws.Range("X2:X" & n).Formula = "=IF(AND(ISNUMBER(SEARCH(""MAYBANK"",$F2)),OR(ISNUMBER(SEARCH(""SINGAPORE"",$F2)),ISNUMBER(SEARCH(""SPORE"",$F2)))),ROUND(O2*0.09,2)+ROUND(P2*0.09,2)+ROUND(Q2*0.09,2),"""")"
    s = n + 1
    ws.Range("J" & s).Formula = "=SUM(J2:J" & n & ")"
    ws.Range("T" & s).Formula = "=SUM(T2:T" & n & ")"
    ws.Range("U" & s).Formula = "=SUM(U2:U" & n & ")"
    ws.Range("V" & s).Formula = "=SUM(V2:V" & n & ")"
    ws.Range("W" & s).Formula = "=SUM(W2:W" & n & ")"
    ws.Range("X" & s).Formula = "=SUM(X2:X" & n & ")"
    ws.Range("V" & (s + 2)).Formula = "=V" & s & "+W" & (s + 2)
    ws.Range("W" & (s + 2)).Formula = "=0.4*W" & s
    ws.Range("I" & (s + 4)).Formula = "=U" & s & "+X" & s
    On Error Resume Next
    ThisWorkbook.Names("MTD_HandOff").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="MTD_HandOff", RefersTo:="='" & ws.Name & "'!" & ws.Range("I" & (s + 4)).Address
End Sub

Attribute VB_Name = "Module_Asia"
Option Explicit

Sub RunAsiaAggregation()
    Dim ws As Worksheet, sh3 As Worksheet, n As Long, pt As PivotTable, src As String, c As Range
    Dim picked As Variant, i As Long, p As String, up As String, nm As String
    Dim wbS As Workbook, already As Boolean, v As Variant, r As Long, h As String
    Dim vCiti As Variant, vWed As Variant, vDma As Variant
    Dim g As Long, cur As Variant

    Set ws = ThisWorkbook.Worksheets("General")
    Set sh3 = ThisWorkbook.Worksheets("Sheet3")

    n = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    ws.Range("T" & (n + 1) & ":W" & (ws.Cells(ws.Rows.Count, "T").End(xlUp).Row + 4)).ClearContents
    ws.Range("G" & (n + 1) & ":H" & (n + 12)).ClearContents

    ws.Range("T2:T" & n).Formula = "=O2"
    ws.Range("U2:U" & n).Formula = "=(T2-W2-X2)*IF(ISNUMBER(SEARCH(""LONG CORRIDOR"",$F2)),0.5,0.6)"
    ws.Range("V2:V" & n).Formula = "=(T2-W2-X2)*IF(ISNUMBER(SEARCH(""LONG CORRIDOR"",$F2)),0.5,0.4)"
    ws.Range("W2:W" & n).Formula = "=IF(ISNA(VLOOKUP(M2,Sheet1!$A$1:$B$4,2,FALSE)),""0"",VLOOKUP(M2,Sheet1!$A$1:$B$4,2,FALSE))"

    cur = Array("AUD", "HKD", "IDR", "JPY", "MYR", "NZD", "PHP", "SGD", "THB")
    For g = 0 To UBound(cur)
        ws.Range("G" & (n + 2 + g)).Value = cur(g)
        ws.Range("H" & (n + 2 + g)).Formula = "=SUMIF(K2:K" & n & ",G" & (n + 2 + g) & ",T2:T" & n & ")"
    Next g

    For Each c In ws.Range("X2:X" & n).Cells
        If Len(CStr(c.Value)) = 0 Then c.Value = 0
    Next c

    src = "General!R1C1:R" & n & "C24"
    For Each pt In sh3.PivotTables
        pt.ChangePivotCache ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=src)
        pt.RefreshTable
    Next pt

    picked = Application.GetOpenFilename("Excel Files (*.xls*),*.xls*", , "Pick the CITI, DMA and Wedbush files", , True)
    If VarType(picked) = vbBoolean Then
        MsgBox "Asia sums and pivot tables refreshed. No source files picked, so the reconciliation figures were not populated.", vbExclamation
        Exit Sub
    End If

    For i = LBound(picked) To UBound(picked)
        p = picked(i)
        nm = Mid(p, InStrRev(p, "\") + 1)
        up = UCase(nm)
        Set wbS = Nothing
        already = False
        On Error Resume Next
        Set wbS = Workbooks(nm)
        On Error GoTo 0
        If wbS Is Nothing Then
            Set wbS = Workbooks.Open(p, ReadOnly:=True)
        Else
            already = True
        End If
        v = Empty
        On Error Resume Next
        v = wbS.Names("MTD_HandOff").RefersToRange.Value
        On Error GoTo 0
        If Not already Then wbS.Close SaveChanges:=False
        Set wbS = Nothing

        If IsNumeric(v) Then
            If InStr(up, "CITI") > 0 Then
                vCiti = v
            ElseIf Left(up, 3) = "DMA" Then
                vDma = v
            ElseIf InStr(up, "WEDBUSH") > 0 Then
                vWed = v
            End If
        End If
    Next i

    For r = 1 To 40
        h = UCase(Trim(CStr(sh3.Cells(r, "H").Value)))
        If Len(h) > 0 Then
            If InStr(h, "US CITI") > 0 Then
                If IsNumeric(vCiti) Then sh3.Cells(r, "I").Value = vCiti
            ElseIf InStr(h, "MAYBANK COMM @ WEDBUSH") > 0 Then
                If InStr(h, "(LT)") > 0 Then
                    If IsNumeric(vDma) Then sh3.Cells(r, "I").Value = vDma
                ElseIf InStr(h, "(") = 0 Then
                    If IsNumeric(vWed) Then sh3.Cells(r, "I").Value = vWed
                End If
            End If
        End If
    Next r

    MsgBox "Asia sums and pivot tables refreshed." & vbCrLf & _
           "CITI: " & IIf(IsNumeric(vCiti), Format(vCiti, "#,##0.00"), "not found") & vbCrLf & _
           "Wedbush: " & IIf(IsNumeric(vWed), Format(vWed, "#,##0.00"), "not found") & vbCrLf & _
           "DMA (LT): " & IIf(IsNumeric(vDma), Format(vDma, "#,##0.00"), "not found"), vbExclamation
End Sub

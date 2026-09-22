Attribute VB_Name = "Module2"
Option Explicit

'=========================================================
' Block Import generation
'
' Sheets used:
'   RawData            - Fidessa trades (cols A:P, P = Instrument)
'   Mapping            - Exchange ID -> Market (cols D:E)
'   Shapes from Client - Client shapes pasted from email (header in row 1)
'
' Main macro:
'   RunAllBlockImport  - builds one Block Import file per
'                        Market / Account / Client
'
' Rules:
'   Account (from RawData col B):
'       ends with "AGENCY n"  -> AC(n-1)   (client numbering is one lower,
'                                 e.g. AGENCY 2 -> AC1, AGENCY 10 -> AC9)
'       no number / AGENCY 1 -> flagged as unknown account
'       ends with "RISK n"    -> Rn (ignored)
'   Shapes with a Trade Account Number starting with "R" are ignored.
'
'   Matching (within the same Client / Market / Account):
'       ISIN + Buy/Sell + Gross price (6dp) + Quantity
'       If the client splits one RawData trade into several shape lines,
'       it counts as matched when the shape quantities add up to the
'       RawData quantity. The block import then books the shape lines.
'
'   Row colours in the Block Import file:
'       Blue = line is in the shapes but missing from RawData
'       Red  = line is in RawData but missing from the shapes
'=========================================================

Private Const SHAPES_SHEET As String = "Shapes from Client"

Private Const FLAG_OK As Long = 0
Private Const FLAG_NO_RAW As Long = 1       ' blue
Private Const FLAG_NO_SHAPE As Long = 2     ' red

' RawData columns
Private Const RC_CODE As Long = 1           ' Counterparty account code -> Cpty
Private Const RC_DESC As Long = 2           ' Counterparty account description
Private Const RC_ISIN As Long = 3
Private Const RC_TS As Long = 6             ' Entered timestamp
Private Const RC_VDATE As Long = 7          ' Value date (yyyymmdd)
Private Const RC_VOL As Long = 8
Private Const RC_GROSS As Long = 9
Private Const RC_SIDE As Long = 11          ' B / S
Private Const RC_EXCH As Long = 15          ' Exchange ID
Private Const RC_INSTR As Long = 16         ' Instrument

' Block Import constants
Private Const BI_INSTR_TYPE As String = "VIEW"
Private Const BI_AVG_PRICE As String = "N"
Private Const BI_ENTITY As String = "MBSG"
Private Const BI_CPTY_TYPE As String = "C"
Private Const BI_CAPACITY As String = "AGENCY"
Private Const BI_TIME_SUFFIX As String = " 15:00:00.000000 +0800"

' Working arrays (indexed by sheet row number)
Private rawStatus() As Long
Private rawQty() As Double
Private shpStatus() As Long
Private shpQty() As Double
Private shpRawRef() As Long

'---------------------------------------------------------
' RunAllBlockImport
'---------------------------------------------------------
Sub RunAllBlockImport()

    Dim wsRaw As Worksheet, wsMap As Worksheet, wsShp As Worksheet

    Set wsRaw = ThisWorkbook.Sheets("RawData")
    Set wsMap = ThisWorkbook.Sheets("Mapping")

    On Error Resume Next
    Set wsShp = ThisWorkbook.Sheets(SHAPES_SHEET)
    On Error GoTo 0

    If wsShp Is Nothing Then
        MsgBox "The sheet '" & SHAPES_SHEET & "' does not exist." & vbCrLf & _
               "Please create it and paste the client shapes in.", _
               vbCritical, "Missing Shapes Sheet"
        Exit Sub
    End If

    ' --- Shapes columns (found by header name) ---
    Dim scAcct As Long, scComp As Long, scTDate As Long, scVDate As Long
    Dim scIsin As Long, scSide As Long, scQty As Long, scPrice As Long, scMkt As Long
    Dim missingCols As String

    scAcct = FindHeader(wsShp, "Trade Account Number", missingCols)
    scComp = FindHeader(wsShp, "Company", missingCols)
    scTDate = FindHeader(wsShp, "Trade Date", missingCols)
    scVDate = FindHeader(wsShp, "Value Date", missingCols)
    scIsin = FindHeader(wsShp, "Instrument ISIN", missingCols)
    scSide = FindHeader(wsShp, "Operation Type", missingCols)
    scQty = FindHeader(wsShp, "Quantity", missingCols)
    scPrice = FindHeader(wsShp, "Price", missingCols)
    scMkt = FindHeader(wsShp, "Market", missingCols)

    If missingCols <> "" Then
        MsgBox "These headers were not found in row 1 of '" & SHAPES_SHEET & "':" & _
               vbCrLf & vbCrLf & missingCols, vbCritical, "Shapes Headers Missing"
        Exit Sub
    End If

    ' --- Exchange ID -> Market from Mapping ---
    Dim exchToMarket As Object
    Set exchToMarket = CreateObject("Scripting.Dictionary")

    Dim r As Long
    r = 2
    Do While wsMap.Cells(r, 4).Value <> ""
        exchToMarket(UCase(Trim(CStr(wsMap.Cells(r, 4).Value)))) = UCase(Trim(CStr(wsMap.Cells(r, 5).Value)))
        r = r + 1
    Loop

    Dim rawLast As Long, shpLast As Long
    rawLast = wsRaw.Cells(wsRaw.Rows.Count, RC_DESC).End(xlUp).Row
    shpLast = wsShp.Cells(wsShp.Rows.Count, scIsin).End(xlUp).Row

    If rawLast < 2 And shpLast < 2 Then
        MsgBox "RawData and '" & SHAPES_SHEET & "' are both empty.", vbExclamation, "Nothing To Do"
        Exit Sub
    End If

    If rawLast < 2 Then rawLast = 1
    If shpLast < 2 Then shpLast = 1

    ReDim rawStatus(1 To rawLast)
    ReDim rawQty(1 To rawLast)
    ReDim shpStatus(1 To shpLast)
    ReDim shpQty(1 To shpLast)
    ReDim shpRawRef(1 To shpLast)

    Dim rawFileKey() As String, shpFileKey() As String
    ReDim rawFileKey(1 To rawLast)
    ReDim shpFileKey(1 To shpLast)

    Dim rawGroups As Object, shpGroups As Object
    Dim isinToInstr As Object, acctToCpty As Object
    Set rawGroups = CreateObject("Scripting.Dictionary")
    Set shpGroups = CreateObject("Scripting.Dictionary")
    Set isinToInstr = CreateObject("Scripting.Dictionary")
    Set acctToCpty = CreateObject("Scripting.Dictionary")

    Dim skipped As String
    Dim clientVal As String, acctVal As String, mktName As String, mktCode As String
    Dim isin As String, sideVal As String, priceKey As String, grpKey As String, fKey As String

    Application.ScreenUpdating = False

    ' --- Read RawData ---
    For r = 2 To rawLast

        rawStatus(r) = -1   ' -1 = not used
        isin = UCase(Trim(CStr(wsRaw.Cells(r, RC_ISIN).Value)))

        If isin <> "" Then

            clientVal = ClientFromDesc(CStr(wsRaw.Cells(r, RC_DESC).Value))
            acctVal = AccountFromDesc(CStr(wsRaw.Cells(r, RC_DESC).Value))

            mktName = ""
            If exchToMarket.Exists(UCase(Trim(CStr(wsRaw.Cells(r, RC_EXCH).Value)))) Then
                mktName = exchToMarket(UCase(Trim(CStr(wsRaw.Cells(r, RC_EXCH).Value))))
            End If
            mktCode = MarketCode(mktName)

            If Left(acctVal, 1) = "R" Then
                ' Risk account - ignored

            ElseIf clientVal = "" Or acctVal = "" Or mktCode = "" Then
                skipped = skipped & "  RawData row " & r & ": " & _
                          IIf(clientVal = "", "unknown client; ", "") & _
                          IIf(acctVal = "", "unknown account; ", "") & _
                          IIf(mktCode = "", "unknown market (" & wsRaw.Cells(r, RC_EXCH).Value & ")", "") & vbCrLf

            Else
                sideVal = UCase(Left(Trim(CStr(wsRaw.Cells(r, RC_SIDE).Value)), 1))
                priceKey = Format(Round(CDbl(wsRaw.Cells(r, RC_GROSS).Value), 6), "0.000000")
                fKey = mktCode & "|" & acctVal & "|" & clientVal
                grpKey = fKey & "|" & isin & "|" & sideVal & "|" & priceKey

                rawStatus(r) = FLAG_NO_SHAPE
                rawQty(r) = CDbl(wsRaw.Cells(r, RC_VOL).Value)
                rawFileKey(r) = fKey
                AddToGroup rawGroups, grpKey, r

                If Not isinToInstr.Exists(isin) And Trim(CStr(wsRaw.Cells(r, RC_INSTR).Value)) <> "" Then
                    isinToInstr(isin) = Trim(CStr(wsRaw.Cells(r, RC_INSTR).Value))
                End If
                If Not acctToCpty.Exists(clientVal & "|" & acctVal) Then
                    acctToCpty(clientVal & "|" & acctVal) = PadCpty(wsRaw.Cells(r, RC_CODE).Value)
                End If
            End If
        End If
    Next r

    ' --- Read Shapes ---
    For r = 2 To shpLast

        shpStatus(r) = -1
        isin = UCase(Trim(CStr(wsShp.Cells(r, scIsin).Value)))

        If isin <> "" Then

            clientVal = ClientFromCompany(CStr(wsShp.Cells(r, scComp).Value))
            acctVal = UCase(Trim(CStr(wsShp.Cells(r, scAcct).Value)))
            mktCode = MarketCode(UCase(Trim(CStr(wsShp.Cells(r, scMkt).Value))))

            If Left(acctVal, 1) = "R" Then
                ' Risk account - ignored

            ElseIf clientVal = "" Or acctVal = "" Or mktCode = "" Then
                skipped = skipped & "  Shapes row " & r & ": " & _
                          IIf(clientVal = "", "unknown company (" & wsShp.Cells(r, scComp).Value & "); ", "") & _
                          IIf(acctVal = "", "no account; ", "") & _
                          IIf(mktCode = "", "unknown market (" & wsShp.Cells(r, scMkt).Value & ")", "") & vbCrLf

            Else
                sideVal = UCase(Left(Trim(CStr(wsShp.Cells(r, scSide).Value)), 1))
                priceKey = Format(Round(CDbl(wsShp.Cells(r, scPrice).Value), 6), "0.000000")
                fKey = mktCode & "|" & acctVal & "|" & clientVal
                grpKey = fKey & "|" & isin & "|" & sideVal & "|" & priceKey

                shpStatus(r) = FLAG_NO_RAW
                shpQty(r) = CDbl(wsShp.Cells(r, scQty).Value)
                shpFileKey(r) = fKey
                AddToGroup shpGroups, grpKey, r
            End If
        End If
    Next r

    ' --- Match ---
    Dim k As Variant
    For Each k In rawGroups.Keys
        If shpGroups.Exists(k) Then
            MatchGroup rawGroups(k), shpGroups(k)
        End If
    Next k

    ' --- Build output rows per file (shapes order first, then red RawData rows) ---
    Dim fileRows As Object
    Set fileRows = CreateObject("Scripting.Dictionary")

    Dim rowArr As Variant
    Dim rr As Long

    For r = 2 To shpLast
        If shpStatus(r) = FLAG_OK Then
            rr = shpRawRef(r)
            rowArr = RowFromRaw(wsRaw, rr, shpQty(r), FLAG_OK)
            AddRow fileRows, shpFileKey(r), rowArr

        ElseIf shpStatus(r) = FLAG_NO_RAW Then
            rowArr = RowFromShape(wsShp, r, scTDate, scVDate, scIsin, scSide, scPrice, _
                                  isinToInstr, acctToCpty, shpFileKey(r))
            AddRow fileRows, shpFileKey(r), rowArr
        End If
    Next r

    For r = 2 To rawLast
        If rawStatus(r) = FLAG_NO_SHAPE Then
            rowArr = RowFromRaw(wsRaw, r, rawQty(r), FLAG_NO_SHAPE)
            AddRow fileRows, rawFileKey(r), rowArr
        End If
    Next r

    If fileRows.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "No block import lines were produced." & vbCrLf & _
               "(Risk accounts are ignored.)" & _
               IIf(skipped <> "", vbCrLf & vbCrLf & "Skipped rows:" & vbCrLf & skipped, ""), _
               vbExclamation, "Nothing Exported"
        Exit Sub
    End If

    ' --- Write files ---
    Dim summary As String
    For Each k In fileRows.Keys
        WriteBlockImport CStr(k), fileRows(k), summary
    Next k

    Application.ScreenUpdating = True

    summary = "Block Import complete!" & vbCrLf & vbCrLf & _
              fileRows.Count & " file(s) saved to:" & vbCrLf & "  " & ThisWorkbook.path & vbCrLf & vbCrLf & _
              summary & vbCrLf & _
              "Blue = in shapes, missing from RawData" & vbCrLf & _
              "Red  = in RawData, missing from shapes"

    If skipped <> "" Then
        summary = summary & vbCrLf & vbCrLf & "Skipped rows:" & vbCrLf & skipped
    End If

    MsgBox summary, vbInformation, "Block Import Complete"

End Sub

'---------------------------------------------------------
' MatchGroup
'
' Rows in one group already share Client / Market / Account /
' ISIN / Side / Price. Step 1 pairs equal quantities. Step 2
' treats what is left as matched if the totals agree (the client
' split one trade into several shape lines).
'---------------------------------------------------------
Private Sub MatchGroup(ByVal rawCol As Collection, ByVal shpCol As Collection)

    Dim rv As Variant, sv As Variant
    Dim rr As Long, sr As Long

    ' Step 1 - exact quantity pairs
    For Each rv In rawCol
        rr = CLng(rv)
        For Each sv In shpCol
            sr = CLng(sv)
            If shpStatus(sr) = FLAG_NO_RAW Then
                If shpQty(sr) = rawQty(rr) Then
                    shpStatus(sr) = FLAG_OK
                    shpRawRef(sr) = rr
                    rawStatus(rr) = FLAG_OK
                    Exit For
                End If
            End If
        Next sv
    Next rv

    ' Step 2 - leftover totals
    Dim sumR As Double, sumS As Double
    Dim cntR As Long, cntS As Long, firstR As Long

    For Each rv In rawCol
        rr = CLng(rv)
        If rawStatus(rr) = FLAG_NO_SHAPE Then
            sumR = sumR + rawQty(rr)
            cntR = cntR + 1
            If firstR = 0 Then firstR = rr
        End If
    Next rv

    For Each sv In shpCol
        sr = CLng(sv)
        If shpStatus(sr) = FLAG_NO_RAW Then
            sumS = sumS + shpQty(sr)
            cntS = cntS + 1
        End If
    Next sv

    If cntR > 0 And cntS > 0 And Abs(sumR - sumS) < 0.000001 Then
        For Each rv In rawCol
            rr = CLng(rv)
            If rawStatus(rr) = FLAG_NO_SHAPE Then rawStatus(rr) = FLAG_OK
        Next rv
        For Each sv In shpCol
            sr = CLng(sv)
            If shpStatus(sr) = FLAG_NO_RAW Then
                shpStatus(sr) = FLAG_OK
                shpRawRef(sr) = firstR
            End If
        Next sv
    End If

End Sub

'---------------------------------------------------------
' Row builders - return Array(12 block import values, flag)
'---------------------------------------------------------
Private Function RowFromRaw(wsRaw As Worksheet, ByVal rr As Long, _
                            ByVal qty As Double, ByVal flag As Long) As Variant

    Dim v(0 To 12) As Variant

    v(0) = CLng(wsRaw.Cells(rr, RC_VDATE).Value)
    v(1) = UCase(Left(Trim(CStr(wsRaw.Cells(rr, RC_SIDE).Value)), 1))
    v(2) = qty
    v(3) = Trim(CStr(wsRaw.Cells(rr, RC_INSTR).Value))
    v(4) = BI_INSTR_TYPE
    v(5) = BI_AVG_PRICE
    v(6) = BI_ENTITY
    v(7) = PadCpty(wsRaw.Cells(rr, RC_CODE).Value)
    v(8) = BI_CPTY_TYPE
    v(9) = BI_CAPACITY
    v(10) = Round(CDbl(wsRaw.Cells(rr, RC_GROSS).Value), 6)
    v(11) = Left(Trim(CStr(wsRaw.Cells(rr, RC_TS).Value)), 8) & BI_TIME_SUFFIX
    v(12) = flag

    RowFromRaw = v

End Function

Private Function RowFromShape(wsShp As Worksheet, ByVal sr As Long, _
                              ByVal scTDate As Long, ByVal scVDate As Long, _
                              ByVal scIsin As Long, ByVal scSide As Long, ByVal scPrice As Long, _
                              isinToInstr As Object, acctToCpty As Object, _
                              ByVal fKey As String) As Variant

    Dim v(0 To 12) As Variant
    Dim parts() As String
    Dim isin As String

    parts = Split(fKey, "|")          ' market | account | client
    isin = UCase(Trim(CStr(wsShp.Cells(sr, scIsin).Value)))

    v(0) = DateToYMD(wsShp.Cells(sr, scVDate).Value)
    v(1) = UCase(Left(Trim(CStr(wsShp.Cells(sr, scSide).Value)), 1))
    v(2) = shpQty(sr)
    v(3) = ""
    If isinToInstr.Exists(isin) Then v(3) = isinToInstr(isin)
    v(4) = BI_INSTR_TYPE
    v(5) = BI_AVG_PRICE
    v(6) = BI_ENTITY
    v(7) = ""
    If acctToCpty.Exists(parts(2) & "|" & parts(1)) Then v(7) = acctToCpty(parts(2) & "|" & parts(1))
    v(8) = BI_CPTY_TYPE
    v(9) = BI_CAPACITY
    v(10) = Round(CDbl(wsShp.Cells(sr, scPrice).Value), 6)
    v(11) = CStr(DateToYMD(wsShp.Cells(sr, scTDate).Value)) & BI_TIME_SUFFIX
    v(12) = FLAG_NO_RAW

    RowFromShape = v

End Function

'---------------------------------------------------------
' WriteBlockImport - creates and saves one file
' File name: <Market>_<Account>_<Client>_BlockImport_<yyyymmdd>.xlsx
'---------------------------------------------------------
Private Function WriteBlockImport(ByVal fKey As String, ByVal lineRows As Collection, _
                                  ByRef summary As String) As String

    Dim parts() As String
    parts = Split(fKey, "|")          ' market | account | client

    Dim headers As Variant
    headers = Array("Value date", "Buy/Sell", "Trade volume", "Instrument", "Instrument type", _
                    "Average Price", "Trading entity ID", "Cpty", "Counterparty type", _
                    "Dealing capacity description", "Gross price", "Entered timestamp")

    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "BlockImport"

    Dim c As Long
    For c = 0 To 11
        ws.Cells(1, c + 1).Value = headers(c)
    Next c

    ' Cpty and Entered timestamp kept as text
    ws.Columns(8).NumberFormat = "@"
    ws.Columns(12).NumberFormat = "@"

    Dim outRow As Long, nBlue As Long, nRed As Long
    Dim v As Variant
    outRow = 2

    For Each v In lineRows
        For c = 0 To 11
            ws.Cells(outRow, c + 1).Value = v(c)
        Next c

        If v(12) = FLAG_NO_RAW Then
            ws.Range(ws.Cells(outRow, 1), ws.Cells(outRow, 12)).Interior.Color = RGB(155, 194, 230)
            nBlue = nBlue + 1
        ElseIf v(12) = FLAG_NO_SHAPE Then
            ws.Range(ws.Cells(outRow, 1), ws.Cells(outRow, 12)).Interior.Color = RGB(255, 124, 128)
            nRed = nRed + 1
        End If

        outRow = outRow + 1
    Next v

    ws.Columns("A:L").AutoFit

    ' Date for file name = Entered timestamp date of the first row
    Dim fileDate As String
    fileDate = Left(CStr(ws.Cells(2, 12).Value), 8)

    Dim fileName As String, fullPath As String
    fileName = parts(0) & "_" & parts(1) & "_" & parts(2) & "_BlockImport_" & fileDate & ".xlsx"
    fullPath = ThisWorkbook.path & "\" & fileName

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False

    summary = summary & "  " & fileName & "  (" & (outRow - 2) & " rows"
    If nBlue > 0 Then summary = summary & ", " & nBlue & " blue"
    If nRed > 0 Then summary = summary & ", " & nRed & " red"
    summary = summary & ")" & vbCrLf

    WriteBlockImport = fullPath

End Function

'---------------------------------------------------------
' Helpers
'---------------------------------------------------------
Private Function AccountFromDesc(ByVal desc As String) As String

    Dim parts() As String
    Dim n As Long
    Dim lastTok As String, prevTok As String

    desc = UCase(Application.WorksheetFunction.Trim(desc))
    If desc = "" Then Exit Function

    parts = Split(desc, " ")
    n = UBound(parts)
    lastTok = parts(n)
    If n >= 1 Then prevTok = parts(n - 1)

    If prevTok = "AGENCY" And IsDigits(lastTok) Then
        ' Client account number is one lower than ours (AGENCY 2 -> AC1)
        If CLng(lastTok) >= 2 Then AccountFromDesc = "AC" & (CLng(lastTok) - 1)
    ElseIf prevTok = "RISK" And IsDigits(lastTok) Then
        AccountFromDesc = "R" & CLng(lastTok)
    ElseIf lastTok = "RISK" Then
        AccountFromDesc = "R"
    End If

End Function

Private Function ClientFromDesc(ByVal desc As String) As String
    desc = UCase(Trim(desc))
    If Left(desc, 15) = "MERRILL LYNCH I" Then
        ClientFromDesc = "MLI"
    ElseIf Left(desc, 4) = "BOFA" Then
        ClientFromDesc = "BAMLSE"
    End If
End Function

Private Function ClientFromCompany(ByVal comp As String) As String
    comp = UCase(Trim(comp))
    If InStr(comp, "BOFA") > 0 Or InStr(comp, "BAML") > 0 Then
        ClientFromCompany = "BAMLSE"
    ElseIf InStr(comp, "MLI") > 0 Or InStr(comp, "MERRILL") > 0 Then
        ClientFromCompany = "MLI"
    End If
End Function

' Add new markets here
Private Function MarketCode(ByVal mkt As String) As String
    mkt = UCase(mkt)
    If InStr(mkt, "INDONESIA") > 0 Then
        MarketCode = "ID"
    ElseIf InStr(mkt, "MALAYSIA") > 0 Then
        MarketCode = "MY"
    End If
End Function

Private Function PadCpty(ByVal code As Variant) As String
    Dim s As String
    s = Trim(CStr(code))
    If Len(s) < 7 And IsDigits(s) Then s = String(7 - Len(s), "0") & s
    PadCpty = s
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    IsDigits = (Len(s) > 0 And Not s Like "*[!0-9]*")
End Function

Private Function DateToYMD(ByVal v As Variant) As Variant
    If IsDate(v) Then
        DateToYMD = CLng(Format(CDate(v), "yyyymmdd"))
    Else
        DateToYMD = Trim(CStr(v))
    End If
End Function

Private Function FindHeader(ws As Worksheet, ByVal headerName As String, _
                            ByRef missing As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If UCase(Trim(CStr(ws.Cells(1, c).Value))) = UCase(headerName) Then
            FindHeader = c
            Exit Function
        End If
    Next c
    missing = missing & "  " & headerName & vbCrLf
End Function

Private Sub AddToGroup(groups As Object, ByVal key As String, ByVal rowNum As Long)
    If Not groups.Exists(key) Then groups.Add key, New Collection
    groups(key).Add rowNum
End Sub

Private Sub AddRow(fileRows As Object, ByVal fKey As String, ByVal rowArr As Variant)
    If Not fileRows.Exists(fKey) Then fileRows.Add fKey, New Collection
    fileRows(fKey).Add rowArr
End Sub

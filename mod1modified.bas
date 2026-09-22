Attribute VB_Name = "Module1"

' EOD Trade Processing
' Required sheets:
'
' RawData - Imported trade data
' Mapping - Account and market mappings
' Helper - Export template and HDR value
' Formatted - Temporary formatted output
' Output - Final export sheet
'
' Main macros:
' GenerateMapping - Build or refresh mappings
' RunEOD - Export one client / market
' RunAll - Export all valid combinations
' ClearRawData - Clear imported source data
'---------------------------------------------------------
' GenerateMapping
'
' Creates or refreshes the Mapping sheet using unique Account Codes and Exchange IDs found in RawData.
' Only ran once at the start of the creation of this file, ignore in normal BAU
'---------------------------------------------------------
Sub GenerateMapping()
    Dim wsRaw       As Worksheet
    Dim wsMap       As Worksheet
    Dim lastRow     As Long
    Dim i           As Long
    Dim acctCode    As Variant
    Dim exchID      As Variant
    Dim desc        As String
    Dim clientVal   As String
    Dim marketVal   As String
    Dim acctDict    As Object
    Dim exchDict    As Object
    Dim mapRow      As Long
    Set wsRaw = ThisWorkbook.Sheets("RawData")
    Set acctDict = CreateObject("Scripting.Dictionary")
    Set exchDict = CreateObject("Scripting.Dictionary")
    ' --- Find or create the Mapping sheet ---
    On Error Resume Next
    Set wsMap = ThisWorkbook.Sheets("Mapping")
    On Error GoTo 0
    If wsMap Is Nothing Then
        Set wsMap = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsMap.Name = "Mapping"
    Else
        ' Warn if Mapping already has data beyond the header row
        If wsMap.Cells(2, 1).Value <> "" Then
            Dim overwrite As Integer
            overwrite = MsgBox("The Mapping sheet already contains data." & vbCrLf & _
                               "Regenerating will wipe all existing mappings." & vbCrLf & vbCrLf & _
                               "Are you sure you want to continue?", _
                               vbYesNo + vbExclamation, "Mapping Sheet Exists")
            If overwrite = vbNo Then Exit Sub
        End If
        wsMap.Cells.ClearContents
    End If
    ' --- Write Mapping sheet headers ---
    wsMap.Cells(1, 1).Value = "Counterparty Account Code"
    wsMap.Cells(1, 2).Value = "Client (MLI / BAMLSE)"
    wsMap.Cells(1, 4).Value = "Exchange ID"
    wsMap.Cells(1, 5).Value = "Market"
    ' Style headers
    With wsMap.Range("A1:B1")
        .Font.Bold = True
        .Interior.Color = RGB(189, 215, 238)
    End With
    With wsMap.Range("D1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(189, 215, 238)
    End With
    ' --- Scan RawData for unique codes and exchange IDs ---
    lastRow = wsRaw.Cells(wsRaw.Rows.count, "B").End(xlUp).Row
    For i = 2 To lastRow
        acctCode = wsRaw.Cells(i, 1).Value
        desc = CStr(wsRaw.Cells(i, 2).Value)
        exchID = wsRaw.Cells(i, 15).Value
        ' --- Collect unique account codes ---
        If acctCode <> "" And Not acctDict.Exists(CStr(acctCode)) Then
            ' Determine MLI or BAMLSE from description
            If Left(UCase(desc), 15) = "MERRILL LYNCH I" Then
                clientVal = "MLI"
            ElseIf Left(UCase(desc), 4) = "BOFA" Then
                clientVal = "BAMLSE"
            Else
                clientVal = "UNKNOWN - PLEASE UPDATE"
            End If
            acctDict.Add CStr(acctCode), clientVal
        End If
        ' --- Collect unique exchange IDs ---
        If exchID <> "" And Not exchDict.Exists(CStr(exchID)) Then
            ' Seed known markets; flag unknowns for manual update
            Select Case UCase(CStr(exchID))
                Case "JKT"
                    marketVal = "Indonesia"
                Case "KLS"
                    marketVal = "Malaysia"
                Case Else
                    marketVal = "UNKNOWN - PLEASE UPDATE"
            End Select
            exchDict.Add CStr(exchID), marketVal
        End If
    Next i
    ' --- Write account code mappings (cols A:B) ---
    mapRow = 2
    Dim acctKey As Variant
    For Each acctKey In acctDict.Keys
        wsMap.Cells(mapRow, 1).Value = acctKey
        wsMap.Cells(mapRow, 2).Value = acctDict(acctKey)
        ' Highlight unknowns in red for easy identification
        If acctDict(acctKey) = "UNKNOWN - PLEASE UPDATE" Then
            wsMap.Cells(mapRow, 2).Interior.Color = RGB(255, 199, 206)
        End If
        mapRow = mapRow + 1
    Next acctKey
    ' --- Write exchange ID mappings (cols D:E) ---
    mapRow = 2
    Dim exchKey As Variant
    For Each exchKey In exchDict.Keys
        wsMap.Cells(mapRow, 4).Value = exchKey
        wsMap.Cells(mapRow, 5).Value = exchDict(exchKey)
        If exchDict(exchKey) = "UNKNOWN - PLEASE UPDATE" Then
            wsMap.Cells(mapRow, 5).Interior.Color = RGB(255, 199, 206)
        End If
        mapRow = mapRow + 1
    Next exchKey
    ' --- Autofit columns ---
    wsMap.Columns("A:B").AutoFit
    wsMap.Columns("D:E").AutoFit
    MsgBox "Mapping sheet generated successfully." & vbCrLf & vbCrLf & _
           "Account codes found : " & acctDict.count & vbCrLf & _
           "Exchange IDs found  : " & exchDict.count & vbCrLf & vbCrLf & _
           "Please review any cells highlighted in red and update manually before running the main macro.", _
           vbInformation, "Generate Mapping Complete"
End Sub

'---------------------------------------------------------
' WriteRawFormulas
'
' Writes the RawData helper formulas (Q:T) from row 2 down to the last
' data row, so they are restored even if someone deletes them.
'   Q = Trade Date   R = Settle Date   S = Net Price   T = tradedate
'---------------------------------------------------------
Private Sub WriteRawFormulas(wsRaw As Worksheet, ByVal lastRow As Long)
    If lastRow < 2 Then lastRow = 2
    With wsRaw
        .Range("Q1:T1").Value = Array("Trade Date", "Settle Date", "Net Price", "tradedate")
        .Range("Q2:Q" & lastRow).Formula = "=DATE(LEFT(T2,4),MID(T2,5,2),RIGHT(T2,2))"
        .Range("R2:R" & lastRow).Formula = "=DATE(LEFT(G2,4),MID(G2,5,2),RIGHT(G2,2))"
        .Range("S2:S" & lastRow).Formula = "=N2/H2"
        .Range("T2:T" & lastRow).Formula = "=LEFT(F2,8)"
        .Range("Q2:R" & lastRow).NumberFormat = "dd/mm/yyyy"
    End With
End Sub

'---------------------------------------------------------
' ValidateMapping
'
' Confirms the Mapping sheet exists and all mappings have been completed.
'---------------------------------------------------------
Private Function ValidateMapping(wsMap As Worksheet) As Boolean
    ValidateMapping = False
    On Error Resume Next
    Set wsMap = ThisWorkbook.Sheets("Mapping")
    On Error GoTo 0
    If wsMap Is Nothing Then
        MsgBox "The Mapping sheet does not exist." & vbCrLf & _
               "Please run the 'Generate Mapping' button first.", _
               vbCritical, "Missing Mapping Sheet"
        Exit Function
    End If
    Dim checkRow As Long
    Dim hasUnknown As Boolean
    hasUnknown = False
    checkRow = 2
    Do While wsMap.Cells(checkRow, 1).Value <> "" Or wsMap.Cells(checkRow, 4).Value <> ""
        If InStr(CStr(wsMap.Cells(checkRow, 2).Value), "UNKNOWN") > 0 Then hasUnknown = True
        If InStr(CStr(wsMap.Cells(checkRow, 5).Value), "UNKNOWN") > 0 Then hasUnknown = True
        checkRow = checkRow + 1
    Loop
    If hasUnknown Then
        MsgBox "The Mapping sheet contains entries marked 'UNKNOWN - PLEASE UPDATE'." & vbCrLf & _
               "Please resolve all unknown mappings before running the macro.", _
               vbCritical, "Incomplete Mapping"
        Exit Function
    End If
    ValidateMapping = True
End Function

'---------------------------------------------------------
' BuildDictionaries
'
' Loads mapping values into dictionaries for fast Account Code and Exchange ID lookups.
'---------------------------------------------------------
Private Sub BuildDictionaries(wsMap As Worksheet, _
                               acctToClient As Object, _
                               exchToMarket As Object)
    Dim dictRow As Long
    dictRow = 2
    Do While wsMap.Cells(dictRow, 1).Value <> ""
        Dim aCode As String
        aCode = CStr(wsMap.Cells(dictRow, 1).Value)
        If Not acctToClient.Exists(aCode) Then
            acctToClient.Add aCode, CStr(wsMap.Cells(dictRow, 2).Value)
        End If
        dictRow = dictRow + 1
    Loop
    dictRow = 2
    Do While wsMap.Cells(dictRow, 4).Value <> ""
        Dim eID As String
        eID = CStr(wsMap.Cells(dictRow, 4).Value)
        If Not exchToMarket.Exists(eID) Then
            exchToMarket.Add eID, CStr(wsMap.Cells(dictRow, 5).Value)
        End If
        dictRow = dictRow + 1
    Loop
End Sub

'---------------------------------------------------------
' CountMatches
'
' Returns the number of trades matching the specified Client and Market.
'---------------------------------------------------------
Private Function CountMatches(wsRaw As Worksheet, _
                               rawLastRow As Long, _
                               clientChoice As String, _
                               marketChoice As String, _
                               acctToClient As Object, _
                               exchToMarket As Object) As Long
    Dim count As Long
    Dim r As Long
    Dim rAcct As String, rExch As String
    Dim rClient As String, rMarket As String
    count = 0
    For r = 2 To rawLastRow
        rAcct = CStr(wsRaw.Cells(r, 1).Value)
        rExch = CStr(wsRaw.Cells(r, 15).Value)
        rClient = ""
        rMarket = ""
        If acctToClient.Exists(rAcct) Then rClient = acctToClient(rAcct)
        If exchToMarket.Exists(rExch) Then rMarket = exchToMarket(rExch)
        If rClient = clientChoice And rMarket = marketChoice Then
            count = count + 1
        End If
    Next r
    CountMatches = count
End Function

'---------------------------------------------------------
' RunEOD
'
' Export a single Client / Market combination.
'---------------------------------------------------------
Sub RunEOD()
    Dim wsMap As Worksheet
    If Not ValidateMapping(wsMap) Then Exit Sub
    Dim frmClient As New ClientForm
    frmClient.Show
    Dim clientChoice As String
    clientChoice = frmClient.SelectedClient
    Unload frmClient
    If clientChoice = "" Then
        MsgBox "No client selected. Macro cancelled.", vbExclamation, "Cancelled"
        Exit Sub
    End If
    Dim mktList() As String
    Dim mktCount As Integer
    mktCount = 0
    Dim mRow As Long
    mRow = 2
    Do While wsMap.Cells(mRow, 4).Value <> ""
        ReDim Preserve mktList(mktCount)
        mktList(mktCount) = CStr(wsMap.Cells(mRow, 5).Value)
        mktCount = mktCount + 1
        mRow = mRow + 1
    Loop
    If mktCount = 0 Then
        MsgBox "No markets found in the Mapping sheet (column E is empty).", _
               vbCritical, "No Markets Found"
        Exit Sub
    End If
    Dim frmMarket As MarketForm
    Set frmMarket = New MarketForm
    frmMarket.MktListData = mktList
    frmMarket.MktCountData = mktCount
    frmMarket.BuildButtons
    frmMarket.Show
    Dim marketChoice As String
    marketChoice = frmMarket.SelectedMarket
    Unload frmMarket
    If marketChoice = "" Then
        MsgBox "No market selected. Macro cancelled.", vbExclamation, "Cancelled"
        Exit Sub
    End If
    Dim acctToClient As Object
    Dim exchToMarket As Object
    Set acctToClient = CreateObject("Scripting.Dictionary")
    Set exchToMarket = CreateObject("Scripting.Dictionary")
    Call BuildDictionaries(wsMap, acctToClient, exchToMarket)
    Dim wsRaw As Worksheet
    Set wsRaw = ThisWorkbook.Sheets("RawData")
    Dim rawLastRow As Long
    rawLastRow = wsRaw.Cells(wsRaw.Rows.count, "B").End(xlUp).Row
    Call WriteRawFormulas(wsRaw, rawLastRow)
    Dim matchCount As Long
    matchCount = CountMatches(wsRaw, rawLastRow, clientChoice, marketChoice, acctToClient, exchToMarket)
    If matchCount = 0 Then
        MsgBox "No trades found for Client: " & clientChoice & "  /  Market: " & marketChoice & vbCrLf & _
               "Please check your selections or the raw data in RawData.", _
               vbExclamation, "No Data Found"
        Exit Sub
    End If
    Dim csvPath As String
    csvPath = ProcessCombination(clientChoice, marketChoice, acctToClient, exchToMarket)
    MsgBox "EOD export complete!" & vbCrLf & vbCrLf & _
           "Client  : " & clientChoice & vbCrLf & _
           "Market  : " & marketChoice & vbCrLf & _
           "Rows    : " & matchCount & vbCrLf & vbCrLf & _
           "File saved to:" & vbCrLf & _
           "  " & csvPath, _
           vbInformation, "Export Complete"
End Sub

'---------------------------------------------------------
' RunAll
'
' Export all Client / Market combinations that have matching trade data.
'---------------------------------------------------------
Sub RunAll()
    Dim wsMap As Worksheet
    If Not ValidateMapping(wsMap) Then Exit Sub
    Dim acctToClient As Object
    Dim exchToMarket As Object
    Set acctToClient = CreateObject("Scripting.Dictionary")
    Set exchToMarket = CreateObject("Scripting.Dictionary")
    Call BuildDictionaries(wsMap, acctToClient, exchToMarket)
    Dim clientList() As String
    Dim marketList() As String
    Dim clientCount As Integer
    Dim marketCount As Integer
    clientCount = 0
    marketCount = 0
    Dim cDict As Object
    Dim mDict As Object
    Set cDict = CreateObject("Scripting.Dictionary")
    Set mDict = CreateObject("Scripting.Dictionary")
    Dim scanRow As Long
    scanRow = 2
    Do While wsMap.Cells(scanRow, 1).Value <> ""
        Dim cVal As String
        cVal = CStr(wsMap.Cells(scanRow, 2).Value)
        If Not cDict.Exists(cVal) Then
            cDict.Add cVal, True
            ReDim Preserve clientList(clientCount)
            clientList(clientCount) = cVal
            clientCount = clientCount + 1
        End If
        scanRow = scanRow + 1
    Loop
    scanRow = 2
    Do While wsMap.Cells(scanRow, 4).Value <> ""
        Dim mVal As String
        mVal = CStr(wsMap.Cells(scanRow, 5).Value)
        If Not mDict.Exists(mVal) Then
            mDict.Add mVal, True
            ReDim Preserve marketList(marketCount)
            marketList(marketCount) = mVal
            marketCount = marketCount + 1
        End If
        scanRow = scanRow + 1
    Loop
    Dim wsRaw As Worksheet
    Set wsRaw = ThisWorkbook.Sheets("RawData")
    Dim rawLastRow As Long
    rawLastRow = wsRaw.Cells(wsRaw.Rows.count, "B").End(xlUp).Row
    Call WriteRawFormulas(wsRaw, rawLastRow)
    Dim c As Integer
    Dim m As Integer
    Dim clientChoice As String
    Dim marketChoice As String
    Dim matchCount As Long
    Dim exportedFiles As String
    Dim skippedCombos As String
    Dim exportCount As Integer
    exportedFiles = ""
    skippedCombos = ""
    exportCount = 0
    For c = 0 To clientCount - 1
        For m = 0 To marketCount - 1
            clientChoice = clientList(c)
            marketChoice = marketList(m)
            matchCount = CountMatches(wsRaw, rawLastRow, clientChoice, marketChoice, acctToClient, exchToMarket)
            If matchCount > 0 Then
                Dim csvPath As String
                csvPath = ProcessCombination(clientChoice, marketChoice, acctToClient, exchToMarket)
                exportedFiles = exportedFiles & "  " & csvPath & vbCrLf
                exportCount = exportCount + 1
            Else
                skippedCombos = skippedCombos & "  " & clientChoice & " / " & marketChoice & " (no data)" & vbCrLf
            End If
        Next m
    Next c

    Dim summary As String
    summary = "RunAll complete!" & vbCrLf & vbCrLf & _
              exportCount & " file(s) exported:" & vbCrLf & exportedFiles
    If skippedCombos <> "" Then
        summary = summary & vbCrLf & "Skipped (no data):" & vbCrLf & skippedCombos
    End If
    MsgBox summary, vbInformation, "Export Complete"
End Sub

'---------------------------------------------------------
' ProcessCombination
'
' Creates the formatted export, builds the output sheet and saves the CSV file.
'---------------------------------------------------------
Private Function ProcessCombination(clientChoice As String, _
                                    marketChoice As String, _
                                    acctToClient As Object, _
                                    exchToMarket As Object) As String
    ProcessCombination = ""
    Dim wsRaw      As Worksheet
    Dim wsFmt      As Worksheet
    Dim wsOut      As Worksheet
    Dim wsTemplate As Worksheet
    Set wsRaw = ThisWorkbook.Sheets("RawData")
    Set wsTemplate = ThisWorkbook.Sheets("Helper")
    Dim rawLastRow As Long
    rawLastRow = wsRaw.Cells(wsRaw.Rows.count, "B").End(xlUp).Row
    ' --- Rebuild Formatted sheet ---
    On Error Resume Next
    Set wsFmt = ThisWorkbook.Sheets("Formatted")
    On Error GoTo 0
    If Not wsFmt Is Nothing Then
        Application.DisplayAlerts = False
        wsFmt.Delete
        Application.DisplayAlerts = True
    End If
    Set wsFmt = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("RawData"))
    wsFmt.Name = "Formatted"
    ' --- Write Helper headers into row 1 now (before data), so column positions are fixed ---
    wsTemplate.Range("A1:S1").Copy
    wsFmt.Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    ' Helper header layout (column number):
    '  A=1  Client
    '  B=2  Instrument (ISIN)
    '  C=3  Instrument Name
    '  D=4  RIC Code
    '  E=5  Trans. Num  (Entered timestamp)
    '  F=6  Trade Date         | populated below from RawData col Q (17)
    '  G=7  Settle Date        | from RawData col R (18)
    '  H=8  Amount/Quantity
    '  I=9  Price (Gross)
    '  J=10 Cash Consideration
    '  K=11 Bk. Off. Sysm      | blank
    '  L=12 Book               | blank
    '  M=13 Ticker Sym.        | blank
    '  N=14 Trans. Type        | BUY/SELL from RawData col K (11)
    '  O=15 Trans. Status      | filled with "O" below
    '  P=16 Settle Currency    | blank
    '  Q=17 Settle Short Code  | blank
    '  R=18 Net Price          | from RawData col S (19)
    '  S=19 Net Cash Consid.   | from RawData col N (14)
    ' RawData source columns (1-based) and their target Formatted columns:
    '   RawData C (3)  | Fmt col 2   (Instrument/ISIN)
    '   RawData D (4)  | Fmt col 3   (Instrument Name)
    '   RawData E (5)  | Fmt col 4   (RIC)
    '   RawData F (6)  | Fmt col 5   (Entered timestamp / Trans.Num)
    '   RawData Q (17) | Fmt col 6   (Trade Date)
    '   RawData R (18) | Fmt col 7   (Settle Date)
    '   RawData H (8)  | Fmt col 8   (Volume / Amount)
    '   RawData I (9)  | Fmt col 9   (Gross Price)
    '   RawData J (10) | Fmt col 10  (Gross Consideration / Cash Consid.)
    '   cols 11-13 intentionally blank
    '   RawData K (11) | Fmt col 14  (Trans. Type  |BUY/SELL)
    '   col 15 = Trans. Status | "O", handled separately
    '   RawData L (12) | Fmt col 16  (Settlement Currency)
    '   col 17 intentionally blank
    '   RawData S (19) | Fmt col 18  (Net Price)
    '   RawData N (14) | Fmt col 19  (Net Cash Consideration)
    ' Note: RawData P (16) = Instrument (pasted from Fidessa) - not exported yet
    ' Map: (RawData source col, Formatted dest col)
    Dim colMap As Variant
    colMap = Array( _
        Array(2, 1), Array(3, 2), Array(4, 3), Array(5, 4), Array(17, 6), Array(18, 7), Array(8, 8), Array(9, 9), Array(10, 10), Array(11, 14), Array(12, 16), Array(19, 18), Array(14, 19))
    ' --- Copy matching rows using direct column mapping ---
    Dim fmtRow As Long
    fmtRow = 2
    Dim r As Long
    Dim rAcct As String, rExch As String
    Dim rClient As String, rMarket As String
    For r = 2 To rawLastRow
        rAcct = CStr(wsRaw.Cells(r, 1).Value)
        rExch = CStr(wsRaw.Cells(r, 15).Value)
        rClient = ""
        rMarket = ""
        If acctToClient.Exists(rAcct) Then rClient = acctToClient(rAcct)
        If exchToMarket.Exists(rExch) Then rMarket = exchToMarket(rExch)
        If rClient = clientChoice And rMarket = marketChoice Then
            Dim pair As Variant
            For Each pair In colMap
                wsFmt.Cells(fmtRow, pair(1)).Value = wsRaw.Cells(r, pair(0)).Value
            Next pair
            fmtRow = fmtRow + 1
        End If
    Next r
    Dim fmtLastRow As Long
    fmtLastRow = fmtRow - 1
    ' --- Date formatting on Trade Date (col F=6) and Settle Date (col G=7) ---
    If fmtLastRow >= 2 Then
        With wsFmt.Range(wsFmt.Cells(2, 6), wsFmt.Cells(fmtLastRow, 6))
            .NumberFormat = "[$-14809]d/m/yyyy;@"
        End With
        With wsFmt.Range(wsFmt.Cells(2, 7), wsFmt.Cells(fmtLastRow, 7))
            .NumberFormat = "[$-14809]d/m/yyyy;@"
        End With

        'Round net Price to 6dp
        Dim rNet As Long
        For rNet = 2 To fmtLastRow
            If IsNumeric(wsFmt.Cells(rNet, 18).Value) Then
                wsFmt.Cells(rNet, 18).Value = Round(wsFmt.Cells(rNet, 18).Value, 6)
            End If
        Next rNet

        wsFmt.Range(wsFmt.Cells(2, 18), wsFmt.Cells(fmtLastRow, 18)).NumberFormat = "0.000000"
    End If
    ' --- Replace B/S with BUY/SELL in Trans. Type column (N=14) only ---
    If fmtLastRow >= 2 Then
        With wsFmt.Range(wsFmt.Cells(2, 14), wsFmt.Cells(fmtLastRow, 14))
            .Replace What:="B", Replacement:="BUY", LookAt:=xlWhole, MatchCase:=True
            .Replace What:="S", Replacement:="SELL", LookAt:=xlWhole, MatchCase:=True
        End With
    End If
    ' --- Trans. Status column (O=15) ? "O" for all data rows ---
    If fmtLastRow >= 2 Then
        wsFmt.Range(wsFmt.Cells(2, 15), wsFmt.Cells(fmtLastRow, 15)).Value = "O"
    End If
    ' --- Autofit Trade/Settle date columns ---
    wsFmt.Columns("F:G").AutoFit
    ' --- HDR stamp in row 1 col A, shift data down ---
    wsFmt.Rows("1:1").Insert Shift:=xlDown
    wsFmt.Cells(1, 1).Value = wsTemplate.Range("A5").Value
    ' --- EOF row ---
    fmtLastRow = wsFmt.Cells(wsFmt.Rows.count, "A").End(xlUp).Row + 1
    wsFmt.Cells(fmtLastRow, 1).Value = "EOF"
    ' --- Build Output sheet ---
    On Error Resume Next
    Set wsOut = ThisWorkbook.Sheets("Output")
    On Error GoTo 0
    If Not wsOut Is Nothing Then
        Application.DisplayAlerts = False
        wsOut.Delete
        Application.DisplayAlerts = True
    End If
    Set wsOut = ThisWorkbook.Sheets.Add(After:=wsFmt)
    wsOut.Name = "Output"
    wsFmt.UsedRange.Copy
    wsOut.Range("A1").PasteSpecial Paste:=xlPasteValues
    wsOut.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    Dim outLastRow As Long
    outLastRow = wsOut.Cells(wsOut.Rows.count, "A").End(xlUp).Row
    ' --- Number formatting for Cash Consideration (J=10) and Net Cash Consideration (S=19) ---
    Dim numFmt As String
    numFmt = "#,##0.00;[Red]#,##0.00"
    If outLastRow >= 3 Then
        With wsOut.Range(wsOut.Cells(3, 10), wsOut.Cells(outLastRow - 1, 10))
            .NumberFormat = numFmt
        End With
        With wsOut.Range(wsOut.Cells(3, 19), wsOut.Cells(outLastRow - 1, 19))
            .NumberFormat = numFmt
        End With
    End If
    ' --- CSV export ---
    Dim todayStr As String
    todayStr = Format(Now(), "YYYYMMDD")
    Dim csvPath As String
    csvPath = ThisWorkbook.Path & "\" & _
              "EOD_Maybank_" & StrConv(marketChoice, vbProperCase) & "_" & clientChoice & "_" & todayStr & ".csv"
    wsOut.Copy
    Dim wbExport As Workbook
    Set wbExport = ActiveWorkbook
    Application.DisplayAlerts = False
    wbExport.SaveAs Filename:=csvPath, FileFormat:=xlCSV
    Application.DisplayAlerts = True
    wbExport.Close SaveChanges:=False
    ProcessCombination = csvPath
End Function

'Clear Raw data button
Sub ClearRawData()
    Dim wsRaw As Worksheet
    Set wsRaw = ThisWorkbook.Sheets("RawData")
    Dim lastRow As Long
    lastRow = wsRaw.Cells(wsRaw.Rows.count, "B").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "RawData is already empty.", vbInformation, "Nothing to Clear"
        Exit Sub
    End If
    wsRaw.Range("A2:P" & lastRow).ClearContents
    If lastRow > 2 Then
        wsRaw.Range("Q3:T" & lastRow).ClearContents
    End If
    MsgBox "RawData cleared successfully", vbInformation, "Done"
End Sub


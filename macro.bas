Option Explicit

'======================================================================
' SGX Corporate Actions Tracker - Universe Growth/Shrink + Sort-Lock
'
' HOW TO INSTALL
' 1. Open the workbook in Excel, press Alt+F11 to open the VBA editor.
' 2. File > Import File... and select this .bas file
'    (or: Insert > Module, then paste this entire file's contents in).
' 3. Save the workbook as a Macro-Enabled Workbook (.xlsm).
' 4. Run SetupSheetProtection ONCE (Alt+F8 > SetupSheetProtection > Run).
'    This turns on AutoFilter and locks out Sort (but not Filter) on
'    every Universe/Buyback/Dividend/Offering/Dashboard sheet, for both
'    STI and ex-STI. That's what stops anyone from physically
'    reordering rows and breaking the row-position links to Universe.
' 5. On each Dashboard sheet, insert a button (Developer tab > Insert >
'    Button) and assign it to RefreshSTITracker (on the STI Dashboard)
'    or RefreshExSTITracker (on the ex-STI Dashboard).
' 6. If you want a password on the protection, set PROTECT_PASSWORD
'    below before running SetupSheetProtection.
' 7. Keep a backup of the workbook before your first run.
'
' WHAT RefreshSTITracker / RefreshExSTITracker DO
' - Compare the current number of tickers in "{prefix} Universe" to
'   how many rows the Buybacks/Dividends/Offerings/Dashboard sheets
'   currently cover.
' - If the universe grew, append new rows (with full formulas) for
'   the additional tickers, in Universe's own order.
' - If the universe shrank, clear the now-unused trailing rows.
' - If the count is unchanged, nothing is touched.
' - Temporarily unprotect the affected sheets to make the change, then
'   re-protect them (sort still off, filter still on) and extend the
'   AutoFilter range to cover any new rows.
' - Keep a hidden ticker snapshot ("TickerLog" sheet) purely so the
'   summary message can tell you which tickers were added/removed
'   since the last run.
'======================================================================

Private Const PROTECT_PASSWORD As String = ""   ' set a password here if you want one

'----------------------------------------------------------------------
' PUBLIC ENTRY POINTS
'----------------------------------------------------------------------
Sub RefreshSTITracker()
    RefreshUniverseTracker "STI"
End Sub

Sub RefreshExSTITracker()
    RefreshUniverseTracker "ex-STI"
End Sub

' Run this once (or any time you want to re-apply the sort lock, e.g.
' after inserting a new sheet or if someone changed sheet protection).
Sub SetupSheetProtection()
    Dim prefixes() As Variant
    prefixes = Array("STI", "ex-STI")
    Dim p As Variant

    For Each p In prefixes
        ApplyFilterAndProtect ThisWorkbook.Sheets(p & " Universe"), "E"
        ApplyFilterAndProtect ThisWorkbook.Sheets(p & " Buybacks"), "J"
        ApplyFilterAndProtect ThisWorkbook.Sheets(p & " Dividends"), "H"
        ApplyFilterAndProtect ThisWorkbook.Sheets(p & " Offerings"), "J"
        ApplyFilterAndProtect ThisWorkbook.Sheets(p & " Dashboard"), "F"
    Next p

    MsgBox "Sorting disabled, filtering enabled, on all Universe/Buyback/Dividend/" & _
           "Offering/Dashboard sheets (STI + ex-STI).", vbInformation
End Sub

'----------------------------------------------------------------------
' CORE REFRESH LOGIC
'----------------------------------------------------------------------
Sub RefreshUniverseTracker(ByVal prefix As String)

    Dim wsUniv As Worksheet, wsBB As Worksheet, wsDiv As Worksheet
    Dim wsOff As Worksheet, wsDash As Worksheet, wsLog As Worksheet
    Dim logCol As Long
    Dim lastUnivRow As Long, newCount As Long, newLastRow As Long
    Dim oldLastRowBB As Long, oldLastRowDiv As Long
    Dim oldLastRowOff As Long, oldLastRowDash As Long
    Dim i As Long
    Dim currDict As Object, prevDict As Object
    Dim addedList As String, removedList As String
    Dim addedCount As Long, removedCount As Long
    Dim v As Variant, k As Variant
    Dim lastLogRow As Long

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsUniv = ThisWorkbook.Sheets(prefix & " Universe")
    Set wsBB = ThisWorkbook.Sheets(prefix & " Buybacks")
    Set wsDiv = ThisWorkbook.Sheets(prefix & " Dividends")
    Set wsOff = ThisWorkbook.Sheets(prefix & " Offerings")
    Set wsDash = ThisWorkbook.Sheets(prefix & " Dashboard")
    Set wsLog = GetOrCreateLogSheet()
    logCol = IIf(prefix = "STI", 1, 2)   ' col A = STI snapshot, col B = ex-STI snapshot

    ' ---- current universe tickers (data starts row 5) ----
    lastUnivRow = wsUniv.Cells(wsUniv.Rows.Count, 1).End(xlUp).Row
    If lastUnivRow < 5 Then lastUnivRow = 4   ' empty-universe edge case
    newCount = lastUnivRow - 4
    If newCount > 0 Then
        newLastRow = 4 + newCount - 1
    Else
        newLastRow = 3   ' nothing to write
    End If

    Set currDict = CreateObject("Scripting.Dictionary")
    For i = 5 To lastUnivRow
        v = wsUniv.Cells(i, 1).Value
        If Len(v) > 0 Then currDict(CStr(v)) = True
    Next i

    ' ---- previous snapshot (informational only, for the summary message) ----
    Set prevDict = CreateObject("Scripting.Dictionary")
    lastLogRow = wsLog.Cells(wsLog.Rows.Count, logCol).End(xlUp).Row
    If lastLogRow >= 1 Then
        For i = 1 To lastLogRow
            v = wsLog.Cells(i, logCol).Value
            If Len(v) > 0 Then prevDict(CStr(v)) = True
        Next i
    End If

    addedCount = 0: removedCount = 0
    For Each k In currDict.Keys
        If Not prevDict.Exists(k) Then
            addedCount = addedCount + 1
            addedList = addedList & k & vbCrLf
        End If
    Next k
    For Each k In prevDict.Keys
        If Not currDict.Exists(k) Then
            removedCount = removedCount + 1
            removedList = removedList & k & vbCrLf
        End If
    Next k

    ' ---- current last-used row on each dependent sheet ----
    oldLastRowBB = GetLastRow(wsBB)
    oldLastRowDiv = GetLastRow(wsDiv)
    oldLastRowOff = GetLastRow(wsOff)
    oldLastRowDash = GetLastRow(wsDash)

    Dim anyChange As Boolean
    anyChange = False
    If newLastRow <> oldLastRowBB Then anyChange = True
    If newLastRow <> oldLastRowDiv Then anyChange = True
    If newLastRow <> oldLastRowOff Then anyChange = True
    If newLastRow <> oldLastRowDash Then anyChange = True

    ' ---- unprotect before editing (protection is what keeps sort locked) ----
    On Error Resume Next
    wsBB.Unprotect PROTECT_PASSWORD
    wsDiv.Unprotect PROTECT_PASSWORD
    wsOff.Unprotect PROTECT_PASSWORD
    wsDash.Unprotect PROTECT_PASSWORD
    On Error GoTo CleanFail

    ' ---- grow or shrink each sheet to match the universe's current size ----
    ResizeTrackerSheet wsBB, oldLastRowBB, newLastRow, prefix, "Buyback"
    ResizeTrackerSheet wsDiv, oldLastRowDiv, newLastRow, prefix, "Dividend"
    ResizeTrackerSheet wsOff, oldLastRowOff, newLastRow, prefix, "Offering"
    ResizeDashboardSheet wsDash, oldLastRowDash, newLastRow, prefix

    ' ---- re-protect and extend the AutoFilter range to the new data extent ----
    ApplyFilterAndProtect wsBB, "J"
    ApplyFilterAndProtect wsDiv, "H"
    ApplyFilterAndProtect wsOff, "J"
    ApplyFilterAndProtect wsDash, "F"

    ' ---- save new snapshot for next time's summary message ----
    wsLog.Columns(logCol).ClearContents
    i = 1
    For Each k In currDict.Keys
        wsLog.Cells(i, logCol).Value = k
        i = i + 1
    Next k

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.CalculateFull

    Dim msg As String
    If Not anyChange Then
        msg = prefix & " universe row count is unchanged (" & newCount & " tickers)." & vbCrLf & _
              "Nothing to add or trim."
    Else
        msg = prefix & " trackers and dashboard resized to match the universe (" & newCount & " tickers)." & vbCrLf & vbCrLf
        If addedCount > 0 Then msg = msg & "Added since last run (" & addedCount & "):" & vbCrLf & addedList & vbCrLf
        If removedCount > 0 Then msg = msg & "Removed since last run (" & removedCount & "):" & vbCrLf & removedList
    End If
    MsgBox msg, vbInformation

    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error refreshing " & prefix & " tracker: " & Err.Description, vbCritical
End Sub

'----------------------------------------------------------------------
' PROTECTION / AUTOFILTER HELPER
'----------------------------------------------------------------------
' Clears any existing AutoFilter, re-applies one over A3:<lastCol><lastRow>
' (headers in row 3, data from row 4), then protects the sheet with
' sorting disallowed but filtering allowed.
Private Sub ApplyFilterAndProtect(ws As Worksheet, lastCol As String)
    On Error Resume Next
    ws.Unprotect PROTECT_PASSWORD
    On Error GoTo 0

    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    Dim dataLastRow As Long
    dataLastRow = GetLastRow(ws)
    If dataLastRow < 4 Then dataLastRow = 4   ' filter range needs at least one data row

    ws.Range("A3:" & lastCol & dataLastRow).AutoFilter

    ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, _
               Scenarios:=True, AllowFiltering:=True, AllowSorting:=False, _
               AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
End Sub

'----------------------------------------------------------------------
' RESIZE HELPERS
'----------------------------------------------------------------------
Private Sub ResizeTrackerSheet(ws As Worksheet, oldLastRow As Long, newLastRow As Long, prefix As String, kind As String)
    Dim r As Long, lastCol As String

    Select Case kind
        Case "Buyback": lastCol = "J"
        Case "Dividend": lastCol = "H"
        Case "Offering": lastCol = "J"
    End Select

    ' universe grew: append new rows with full formulas
    If newLastRow > oldLastRow Then
        For r = oldLastRow + 1 To newLastRow
            Select Case kind
                Case "Buyback": WriteBuybackRow ws, r, prefix
                Case "Dividend": WriteDividendRow ws, r, prefix
                Case "Offering": WriteOfferingRow ws, r, prefix
            End Select
        Next r
    End If

    ' universe shrank: clear the now-unused trailing rows
    If oldLastRow > newLastRow Then
        ClearRows ws, newLastRow + 1, oldLastRow, "A", lastCol
    End If
End Sub

Private Sub ResizeDashboardSheet(ws As Worksheet, oldLastRow As Long, newLastRow As Long, prefix As String)
    Dim r As Long

    If newLastRow > oldLastRow Then
        For r = oldLastRow + 1 To newLastRow
            WriteDashboardRow ws, r, prefix
        Next r
    End If

    If oldLastRow > newLastRow Then
        ClearRows ws, newLastRow + 1, oldLastRow, "A", "F"
    End If
End Sub

'----------------------------------------------------------------------
' MISC HELPERS
'----------------------------------------------------------------------
Private Function GetOrCreateLogSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("TickerLog")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "TickerLog"
        ws.Visible = xlSheetVeryHidden
    End If
    Set GetOrCreateLogSheet = ws
End Function

Private Function GetLastRow(ws As Worksheet) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If r < 4 Then r = 3
    GetLastRow = r
End Function

Private Sub ClearRows(ws As Worksheet, startRow As Long, endRow As Long, colFrom As String, colTo As String)
    If endRow < startRow Then Exit Sub
    Dim c1 As Long, c2 As Long
    c1 = ws.Range(colFrom & "1").Column
    c2 = ws.Range(colTo & "1").Column
    ws.Range(ws.Cells(startRow, c1), ws.Cells(endRow, c2)).ClearContents
End Sub

'----------------------------------------------------------------------
' FULL-ROW WRITERS (used only for brand-new rows when universe grows)
'----------------------------------------------------------------------
Private Sub WriteBuybackRow(ws As Worksheet, r As Long, prefix As String)
    Dim u As Long: u = r + 1
    ws.Cells(r, 1).Formula = "='" & prefix & " Universe'!A" & u
    ws.Cells(r, 2).Formula = "=IFERROR(_xll.BDP(A" & r & ",""NAME""),"""")"
    ws.Cells(r, 3).Formula = "=IFERROR(_xll.BDP(A" & r & ",""EQY_SH_OUT""),NA())"
    ws.Cells(r, 4).Formula = "=IFERROR(_xll.BDH(A" & r & ",""EQY_SH_OUT"",""-""&Settings!$B$4&""AW"",""-""&Settings!$B$4&""AW"",""days=a"",""fill=p""),NA())"
    ws.Cells(r, 5).Formula = "=IFERROR((C" & r & "-D" & r & ")/D" & r & ",NA())"
    ws.Cells(r, 6).Formula = "=IF(ISNUMBER(E" & r & "),IF(E" & r & "<=Settings!$B$6,""BUYBACK ACTIVITY"",""""),"""")"
    ws.Cells(r, 7).Formula = "=IFERROR(_xll.BDP(A" & r & ",""VOLUME_AVG_30D""),NA())"
    ws.Cells(r, 8).Formula = "=IFERROR(_xll.BDP(A" & r & ",""PX_VOLUME""),NA())"
    ws.Cells(r, 9).Formula = "=IFERROR(H" & r & "/G" & r & ",NA())"
    ws.Cells(r, 10).Formula = "=IF(ISNUMBER(I" & r & "),IF(I" & r & ">=Settings!$B$7,""VOL SPIKE"",""""),"""")"
End Sub

Private Sub WriteDividendRow(ws As Worksheet, r As Long, prefix As String)
    Dim u As Long: u = r + 1
    ws.Cells(r, 1).Formula = "='" & prefix & " Universe'!A" & u
    ws.Cells(r, 2).Formula = "=IFERROR(_xll.BDP(A" & r & ",""NAME""),"""")"
    ws.Cells(r, 3).Formula = "=IF(LEFT((_xll.BDP(A" & r & ",""DVD_DECLARED_DT"")),4)=""#N/A"",NA(),(_xll.BDP(A" & r & ",""DVD_DECLARED_DT"")))"
    ws.Cells(r, 4).Formula = "=IF(LEFT((_xll.BDP(A" & r & ",""DVD_EX_DT"")),4)=""#N/A"",NA(),(_xll.BDP(A" & r & ",""DVD_EX_DT"")))"
    ws.Cells(r, 5).Formula = "=IF(LEFT((_xll.BDP(A" & r & ",""DVD_SH_LAST"")), 4) = ""#N/A"", NA(), (_xll.BDP(A" & r & ",""DVD_SH_LAST"")))"
    ws.Cells(r, 6).Formula = "=IF(LEFT((_xll.BDP(A" & r & ",""DVD_TYP_LAST"")), 4) = ""#N/A"", NA(), (_xll.BDP(A" & r & ",""DVD_TYP_LAST"")))"
    ws.Cells(r, 7).Formula = "=IF(ISNUMBER(C" & r & "),IF(C" & r & ">=Settings!$B$5,""Y"",""N""),"""")"
    ws.Cells(r, 8).Formula = "=IF(G" & r & "=""Y"",IF(ISNUMBER(SEARCH(""Special"",F" & r & ")),""SPECIAL DIVIDEND"",""NEW DIVIDEND""),IF(ISNUMBER(SEARCH(""Special"",F" & r & ")),""SPECIAL DIVIDEND"",""""))"
End Sub

Private Sub WriteOfferingRow(ws As Worksheet, r As Long, prefix As String)
    Dim u As Long: u = r + 1
    ws.Cells(r, 1).Formula = "='" & prefix & " Universe'!A" & u
    ws.Cells(r, 2).Formula = "=IFERROR(_xll.BDP(A" & r & ",""NAME""),"""")"
    ws.Range(ws.Cells(r, 3), ws.Cells(r, 8)).FormulaArray = _
        "=_xll.BDS(A" & r & ",""EQUITY_OFFERINGS"",""sortdesc=3"",""endrow=1"")"
    ws.Cells(r, 9).Formula = "=IF(ISNUMBER(E" & r & "),IF(E" & r & ">=Settings!$B$5,""Y"",""N""),"""")"
    ws.Cells(r, 10).Formula = "=IF(I" & r & "=""Y"",""RECENT OFFERING/SPLIT/RIGHTS"","""")"
End Sub

Private Sub WriteDashboardRow(ws As Worksheet, r As Long, prefix As String)
    Dim u As Long: u = r + 1
    ws.Cells(r, 1).Formula = "='" & prefix & " Universe'!A" & u
    ws.Cells(r, 2).Formula = "='" & prefix & " Universe'!B" & u
    ws.Cells(r, 3).Formula = "='" & prefix & " Buybacks'!F" & r
    ws.Cells(r, 4).Formula = "='" & prefix & " Buybacks'!J" & r
    ws.Cells(r, 5).Formula = "='" & prefix & " Dividends'!H" & r
    ws.Cells(r, 6).Formula = "='" & prefix & " Offerings'!J" & r
End Sub

'======================================================================
' SORTED VIEW SHEETS (no protection-lifting needed to "sort")
'
' Run CreateSortedViews() once to build six extra sheets:
'   STI Buybacks (Sorted), STI Dividends (Sorted), STI Offerings (Sorted)
'   ex-STI Buybacks (Sorted), ex-STI Dividends (Sorted), ex-STI Offerings (Sorted)
'
' Each one mirrors its source tracker sheet live via SORT(FILTER(...)),
' controlled by two yellow input cells (B2 = column number to sort by,
' D2 = 1 for ascending / -1 for descending). Changing those two cells
' re-sorts instantly - nothing about the source sheet is touched, so
' there is nothing for the refresh macro to ever need to undo here.
' Re-run CreateSortedViews() any time (e.g. after SetupSheetProtection)
' to rebuild these from scratch if you want a clean copy.
'======================================================================

Sub CreateSortedViews()
    BuildSortedView "STI", "Buybacks", "J"
    BuildSortedView "STI", "Dividends", "H"
    BuildSortedView "STI", "Offerings", "J"
    BuildSortedView "ex-STI", "Buybacks", "J"
    BuildSortedView "ex-STI", "Dividends", "H"
    BuildSortedView "ex-STI", "Offerings", "J"
    MsgBox "Sorted-view sheets created for all six trackers.", vbInformation
End Sub

Private Sub BuildSortedView(ByVal prefix As String, ByVal kind As String, ByVal lastColLetter As String)
    Dim srcName As String, viewName As String
    Dim ws As Worksheet
    Dim lastColNum As Long, c As Long

    srcName = prefix & " " & kind
    viewName = Left(prefix & " " & kind & " (Sorted)", 31)

    ' rebuild fresh each time
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(viewName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = viewName
    lastColNum = ws.Range(lastColLetter & "1").Column

    ' ---- controls ----
    ws.Range("A1").Value = srcName & " - Sorted View (read-only; edit the source sheet's data, not this one)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Sort by column #:"
    ws.Range("B2").Value = 1
    ws.Range("B2").Interior.Color = RGB(255, 255, 0)
    ws.Range("C2").Value = "Order (1 = A-Z, -1 = Z-A):"
    ws.Range("D2").Value = 1
    ws.Range("D2").Interior.Color = RGB(255, 255, 0)

    ' ---- header row 4, linked live to the source's header row 3 ----
    For c = 1 To lastColNum
        ws.Cells(4, c).Formula = "='" & srcName & "'!" & ColLetter(c) & "3"
    Next c
    ws.Rows(4).Font.Bold = True

    ' ---- spilling sort formula, anchored at A5 ----
    Dim rangeRef As String, condRef As String
    rangeRef = "'" & srcName & "'!A4:" & lastColLetter & "10000"
    condRef = "'" & srcName & "'!A4:A10000"
    ws.Range("A5").Formula2 = "=IFERROR(SORT(FILTER(" & rangeRef & "," & condRef & "<>""""),$B$2,$D$2),""No data"")"

    ws.Columns("A:" & lastColLetter).AutoFit
    ws.Tab.Color = RGB(198, 224, 255)

    ' ---- lock everything except the two input cells ----
    ws.Cells.Locked = True
    ws.Range("B2").Locked = False
    ws.Range("D2").Locked = False
    ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, _
               Scenarios:=True, AllowFiltering:=True
End Sub

Private Function ColLetter(ByVal n As Long) As String
    Dim addr As String
    addr = Cells(1, n).Address(False, False)
    ColLetter = Left(addr, Len(addr) - 1)
End Function

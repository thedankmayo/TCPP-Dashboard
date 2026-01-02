Attribute VB_Name = "modTCPPv2"
Option Explicit

'========================
' v2: UserForm-driven hub
'========================

'--- Sheet names
Private Const SH_HOME As String = "HOME"
Private Const SH_LEDGER As String = "DATA_Ledger"
Private Const SH_LOOKUPS As String = "DATA_Lookups"
Private Const SH_BUDGET As String = "DATA_Budget"
Private Const SH_MONTHSTATUS As String = "DATA_MonthStatus"
Private Const SH_AUDIT As String = "DATA_Audit"
Private Const SH_REPORT As String = "RPT_Monthly"

'--- Table names
Private Const T_LEDGER As String = "tblLedger"
Private Const T_COA As String = "tblCOA"
Private Const T_EVENTS As String = "tblEvents"
Private Const T_CHARITIES As String = "tblCharities"
Private Const T_PAYMETHOD As String = "tblPaymentMethods"
Private Const T_CONFIG As String = "tblConfig"
Private Const T_BUDGET As String = "tblBudget"
Private Const T_MONTHSTATUS As String = "tblMonthStatus"
Private Const T_AUDIT As String = "tblAuditLog"

'--- Config keys
Private Const CFG_FISCAL_START_MONTH As String = "FiscalYearStartMonth"
Private Const CFG_APPROVER_NAME As String = "ApproverName"
Private Const CFG_RECEIPT_THRESHOLD As String = "ReceiptRequiredThreshold"
Private Const CFG_RECEIPTS_FOLDER As String = "ReceiptsFolderRelative"
Private Const CFG_BOARDPACKETS_FOLDER As String = "BoardPacketsFolderRelative"
Private Const CFG_LOCKS_ENABLED As String = "CloseLocksEnabled"

'--- Globals for current dashboard filters
Public gMonthKey As String
Public gEventFilter As String
Public gCharityFilter As String

'========================
' Public entry points
'========================

Public Sub InitializeTool(ByVal forceRebuild As Boolean)
    Dim prevSU As Boolean, prevEV As Boolean
    prevSU = Application.ScreenUpdating
    prevEV = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    EnsureCoreSheets forceRebuild
    EnsureCoreTables forceRebuild
    SeedLookupsIfEmpty
    EnsureConfigDefaults
    LockDownWorkbookUI

    ' default filters
    If Len(gMonthKey) = 0 Then gMonthKey = Format(Date, "yyyy-mm")
    If Len(gEventFilter) = 0 Then gEventFilter = "(All)"
    If Len(gCharityFilter) = 0 Then gCharityFilter = "(All)"

    Application.EnableEvents = prevEV
    Application.ScreenUpdating = prevSU
End Sub

Public Sub ShowDashboard()
    On Error Resume Next
    Unload frmDashboard
    On Error GoTo 0
    frmDashboard.Show vbModal
End Sub

Public Sub CleanupOnClose()
    On Error Resume Next
    Application.DisplayAlerts = True
End Sub

Public Sub RunSelfTest()
    Dim msg As String
    msg = SelfTestReport()
    MsgBox msg, vbInformation, "TCPP v2 Self Test"
End Sub

'========================
' Core structure
'========================

Private Sub EnsureCoreSheets(ByVal forceRebuild As Boolean)
    EnsureSheet SH_HOME, xlSheetVisible
    EnsureSheet SH_LEDGER, xlSheetVeryHidden
    EnsureSheet SH_LOOKUPS, xlSheetVeryHidden
    EnsureSheet SH_BUDGET, xlSheetVeryHidden
    EnsureSheet SH_MONTHSTATUS, xlSheetVeryHidden
    EnsureSheet SH_AUDIT, xlSheetVeryHidden
    EnsureSheet SH_REPORT, xlSheetVeryHidden

    ' keep HOME minimal if newly created
    With GetSheet(SH_HOME)
        .Cells.ClearContents
        .Range("A1").value = "TCPP Treasurer Dashboard (v2)"
        .Range("A2").value = "UserForm hub. Keep this sheet open."
    End With
End Sub

Private Sub EnsureCoreTables(ByVal forceRebuild As Boolean)
    EnsureLookupTables forceRebuild
    EnsureLedgerTable forceRebuild
    EnsureBudgetTable forceRebuild
    EnsureMonthStatusTable forceRebuild
    EnsureAuditTable forceRebuild
    EnsureReportSheetLayout
End Sub

Private Sub EnsureLookupTables(ByVal forceRebuild As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_LOOKUPS)

    EnsureTable ws, T_COA, Array("Category"), 1, 1, forceRebuild
    EnsureTable ws, T_EVENTS, Array("Event"), 1, 3, forceRebuild
    EnsureTable ws, T_CHARITIES, Array("Charity"), 1, 5, forceRebuild
    EnsureTable ws, T_PAYMETHOD, Array("PaymentMethod"), 1, 7, forceRebuild
    EnsureTable ws, T_CONFIG, Array("Key", "Value"), 1, 9, forceRebuild
End Sub

Private Sub EnsureLedgerTable(ByVal forceRebuild As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_LEDGER)

    Dim headers As Variant
    headers = Array( _
        "TxnID", "Date", "MonthKey", "FiscalYear", _
        "TxnType", "TxnDetail", _
        "Category", "Event", "Charity", _
        "Gross", "Fees", "Net", _
        "PaymentMethod", "PayeeOrSource", "Memo", _
        "ReceiptRequired", "ReceiptStatus", "ReceiptLink", _
        "ReceiptWaivedReason", "ReceiptWaivedBy", "ReceiptWaivedAt", _
        "ApprovedBy", "ClosedFlag", _
        "CreatedAt", "UpdatedAt" _
    )

    EnsureTable ws, T_LEDGER, headers, 1, 1, forceRebuild
End Sub

Private Sub EnsureBudgetTable(ByVal forceRebuild As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_BUDGET)
    EnsureTable ws, T_BUDGET, Array("MonthKey", "FiscalYear", "Category", "BudgetAmount"), 1, 1, forceRebuild
End Sub

Private Sub EnsureMonthStatusTable(ByVal forceRebuild As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_MONTHSTATUS)
    EnsureTable ws, T_MONTHSTATUS, Array( _
        "MonthKey", "FiscalYear", _
        "BeginningBalance", "EndingBalance", _
        "LedgerDeposits", "LedgerWithdrawals", _
        "ExpectedEnding", "ReconDifference", _
        "ReconStatus", "LastReconAt", _
        "ClosedFlag", "ClosedAt" _
    ), 1, 1, forceRebuild
End Sub

Private Sub EnsureAuditTable(ByVal forceRebuild As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_AUDIT)
    EnsureTable ws, T_AUDIT, Array("Timestamp", "User", "Action", "TxnID", "Details"), 1, 1, forceRebuild
End Sub

Private Sub EnsureReportSheetLayout()
    Dim ws As Worksheet
    Set ws = GetSheet(SH_REPORT)
    ws.Cells.ClearContents

    ws.Range("A1").value = "TCPP Board Packet (Monthly)"
    ws.Range("A3").value = "Month"
    ws.Range("B3").value = ""
    ws.Range("A4").value = "Fiscal Year"
    ws.Range("B4").value = ""

    ws.Range("A6").value = "Summary (Month Only)"
    ws.Range("A7").value = "Beginning Cash"
    ws.Range("A8").value = "Total Income (Net)"
    ws.Range("A9").value = "Total Expenses (Net)"
    ws.Range("A10").value = "Net Change"
    ws.Range("A11").value = "Ending Cash"

    ws.Range("D6").value = "Controls"
    ws.Range("D7").value = "Reconciled?"
    ws.Range("D8").value = "Closed?"
    ws.Range("D9").value = "Missing Receipts (count / $)"
    ws.Range("D10").value = "Uncategorized (count)"

    ws.Range("A13").value = "Budget vs Actual (Month)"
    ws.Range("A14").value = "Category"
    ws.Range("B14").value = "Budget"
    ws.Range("C14").value = "Actual"
    ws.Range("D14").value = "Var $"
    ws.Range("E14").value = "Var %"

    ws.Range("G13").value = "Charity (Month + YTD)"
    ws.Range("G14").value = "Raised (Net)"
    ws.Range("G15").value = "Paid Out (Net)"
    ws.Range("G16").value = "Held (YTD Net)"

    ws.Range("A24").value = "Event Rollup (Month)"
    ws.Range("A25").value = "Event"
    ws.Range("B25").value = "Income (Net)"
    ws.Range("C25").value = "Expenses (Net)"
    ws.Range("D25").value = "Net"

    ws.Range("A35").value = "YTD (Jun ? Selected Month)"
    ws.Range("A36").value = "Income (Net)"
    ws.Range("A37").value = "Expenses (Net)"
    ws.Range("A38").value = "Net"
    ws.Range("A39").value = "Budget Var (YTD $)"
End Sub

Private Sub EnsureConfigDefaults()
    Dim cfg As ListObject
    Set cfg = GetTable(SH_LOOKUPS, T_CONFIG)

    UpsertConfig cfg, CFG_FISCAL_START_MONTH, "6"
    UpsertConfig cfg, CFG_APPROVER_NAME, Application.UserName
    UpsertConfig cfg, CFG_RECEIPT_THRESHOLD, "0"
    UpsertConfig cfg, CFG_RECEIPTS_FOLDER, ".\Receipts\"
    UpsertConfig cfg, CFG_BOARDPACKETS_FOLDER, ".\BoardPackets\"
    UpsertConfig cfg, CFG_LOCKS_ENABLED, "TRUE"
End Sub

Private Sub SeedLookupsIfEmpty()
    ' COA categories
    Dim coa As ListObject: Set coa = GetTable(SH_LOOKUPS, T_COA)
    If coa.DataBodyRange Is Nothing Then
        AppendListValue coa, 1, "Administrative"
        AppendListValue coa, 1, "Programs"
        AppendListValue coa, 1, "Fundraising"
        AppendListValue coa, 1, "Marketing"
        AppendListValue coa, 1, "Travel"
        AppendListValue coa, 1, "Services"
        AppendListValue coa, 1, "Misc"
    End If

    Dim ev As ListObject: Set ev = GetTable(SH_LOOKUPS, T_EVENTS)
    If ev.DataBodyRange Is Nothing Then
        AppendListValue ev, 1, "Backrooms"
        AppendListValue ev, 1, "Tank"
        AppendListValue ev, 1, "Class101"
        AppendListValue ev, 1, "Class201"
        AppendListValue ev, 1, "Class301"
        AppendListValue ev, 1, "Walkies"
        AppendListValue ev, 1, "Barkade & Pupparoni (B&P)"
    End If

    Dim ch As ListObject: Set ch = GetTable(SH_LOOKUPS, T_CHARITIES)
    If ch.DataBodyRange Is Nothing Then
        AppendListValue ch, 1, "Aliveness Project"
    End If

    Dim pm As ListObject: Set pm = GetTable(SH_LOOKUPS, T_PAYMETHOD)
    If pm.DataBodyRange Is Nothing Then
        AppendListValue pm, 1, "Cash"
        AppendListValue pm, 1, "Card"
        AppendListValue pm, 1, "Zeffy"
        AppendListValue pm, 1, "Bank"
        AppendListValue pm, 1, "Other"
    End If
End Sub

Private Sub LockDownWorkbookUI()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(SH_HOME).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(SH_HOME).Activate
    ActiveWindow.DisplayWorkbookTabs = False
    Application.DisplayAlerts = True
End Sub

'========================
' Sheet/table helpers
'========================

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
End Function

Private Sub EnsureSheet(ByVal sheetName As String, ByVal vis As XlSheetVisibility)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = sheetName
    End If
    ws.Visible = vis
End Sub

Private Function GetTable(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Dim ws As Worksheet: Set ws = GetSheet(sheetName)
    Set GetTable = ws.ListObjects(tableName)
End Function

Private Sub EnsureTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal headers As Variant, _
                        ByVal startRow As Long, ByVal startCol As Long, ByVal forceRebuild As Boolean)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    If Not lo Is Nothing Then
        If forceRebuild Then
            lo.Unlist
            Set lo = Nothing
        Else
            ' ensure headers match (add any missing headers at end)
            EnsureHeaders lo, headers
            Exit Sub
        End If
    End If

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(startRow, startCol + i).value = headers(i)
        ws.Cells(startRow + 1, startCol + i).value = ""
    Next i

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 1, startCol + UBound(headers)))

    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo.name = tableName

    ' remove the placeholder blank row
    If Not lo.DataBodyRange Is Nothing Then
        If lo.ListRows.count = 1 Then lo.ListRows(1).Delete
    End If
End Sub

Private Sub EnsureHeaders(ByVal lo As ListObject, ByVal headers As Variant)
    Dim existing As Object: Set existing = CreateObject("Scripting.Dictionary")
    existing.CompareMode = 1

    Dim i As Long
    For i = 1 To lo.ListColumns.count
        existing(lo.ListColumns(i).name) = True
    Next i

    For i = LBound(headers) To UBound(headers)
        If Not existing.Exists(CStr(headers(i))) Then
            lo.ListColumns.Add.name = CStr(headers(i))
        End If
    Next i
End Sub

Private Sub AppendListValue(ByVal lo As ListObject, ByVal colIndex As Long, ByVal value As String)
    Dim lr As ListRow
    Set lr = lo.ListRows.Add
    lr.Range.Cells(1, colIndex).value = value
End Sub

Private Sub UpsertConfig(ByVal lo As ListObject, ByVal key As String, ByVal value As String)
    Dim r As Range
    If lo.DataBodyRange Is Nothing Then
        AppendConfig lo, key, value
        Exit Sub
    End If

    For Each r In lo.ListColumns(1).DataBodyRange.Cells
        If CStr(r.value) = key Then
            r.Offset(0, 1).value = value
            Exit Sub
        End If
    Next r
    AppendConfig lo, key, value
End Sub

Private Sub AppendConfig(ByVal lo As ListObject, ByVal key As String, ByVal value As String)
    Dim lr As ListRow: Set lr = lo.ListRows.Add
    lr.Range.Cells(1, 1).value = key
    lr.Range.Cells(1, 2).value = value
End Sub

Private Function GetConfigValue(ByVal key As String, Optional ByVal defaultValue As String = "") As String
    Dim cfg As ListObject: Set cfg = GetTable(SH_LOOKUPS, T_CONFIG)
    If cfg.DataBodyRange Is Nothing Then
        GetConfigValue = defaultValue
        Exit Function
    End If

    Dim r As Range
    For Each r In cfg.ListColumns(1).DataBodyRange.Cells
        If CStr(r.value) = key Then
            GetConfigValue = CStr(r.Offset(0, 1).value)
            Exit Function
        End If
    Next r
    GetConfigValue = defaultValue
End Function

'========================
' Fiscal/month helpers
'========================

Public Function FiscalYearForMonthKey(ByVal monthKey As String) As Long
    Dim y As Long, m As Long
    y = CLng(Left$(monthKey, 4))
    m = CLng(Right$(monthKey, 2))

    Dim startM As Long
    startM = CLng(GetConfigValue(CFG_FISCAL_START_MONTH, "6"))

    If m >= startM Then
        FiscalYearForMonthKey = y + 1
    Else
        FiscalYearForMonthKey = y
    End If
End Function

Public Function MonthKeyFromDate(ByVal d As Date) As String
    MonthKeyFromDate = Format(d, "yyyy-mm")
End Function

Private Function ParseDateFlexible(ByVal s As String) As Date
    If IsDate(s) Then
        ParseDateFlexible = CDate(s)
    Else
        ParseDateFlexible = Date
    End If
End Function

Private Function NzStr(ByVal v As Variant, Optional ByVal fallback As String = "") As String
    If IsError(v) Then
        NzStr = fallback
    ElseIf IsNull(v) Then
        NzStr = fallback
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        NzStr = fallback
    Else
        NzStr = CStr(v)
    End If
End Function

Private Function NzDbl(ByVal v As Variant, Optional ByVal fallback As Double = 0#) As Double
    If IsError(v) Or IsNull(v) Or Len(Trim$(CStr(v))) = 0 Then
        NzDbl = fallback
    Else
        NzDbl = CDbl(v)
    End If
End Function

'========================
' Audit
'========================

Public Sub AuditLog(ByVal action As String, Optional ByVal txnId As String = "", Optional ByVal details As String = "")
    Dim lo As ListObject: Set lo = GetTable(SH_AUDIT, T_AUDIT)
    Dim lr As ListRow: Set lr = lo.ListRows.Add
    lr.Range.Cells(1, 1).value = Now
    lr.Range.Cells(1, 2).value = Application.UserName
    lr.Range.Cells(1, 3).value = action
    lr.Range.Cells(1, 4).value = txnId
    lr.Range.Cells(1, 5).value = details
End Sub

'========================
' Ledger operations
'========================

Public Function NextTxnId(ByVal monthKey As String) As String
    Dim fy As Long: fy = FiscalYearForMonthKey(monthKey)

    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    Dim maxN As Long: maxN = 0

    If Not lo.DataBodyRange Is Nothing Then
        Dim r As Range
        For Each r In lo.ListColumns("TxnID").DataBodyRange.Cells
            Dim id As String: id = CStr(r.value)
            If InStr(1, id, "TCPP-", vbTextCompare) = 1 Then
                Dim parts() As String: parts = Split(id, "-")
                If UBound(parts) >= 3 Then
                    If CLng(parts(1)) = fy Then
                        Dim seq As Long: seq = CLng(parts(3))
                        If seq > maxN Then maxN = seq
                    End If
                End If
            End If
        Next r
    End If

    NextTxnId = "TCPP-" & CStr(fy) & "-" & Replace(monthKey, "-", "") & "-" & Format$(maxN + 1, "0000")
End Function

Public Function AddLedgerEntry( _
    ByVal txnDate As Date, ByVal txnType As String, ByVal txnDetail As String, _
    ByVal category As String, ByVal eventName As String, ByVal charityName As String, _
    ByVal gross As Double, ByVal fees As Double, ByVal paymentMethod As String, _
    ByVal payeeOrSource As String, ByVal memo As String, _
    ByVal receiptRequired As Boolean, _
    Optional ByVal allowInClosedMonth As Boolean = False _
) As String

    Dim monthKey As String: monthKey = MonthKeyFromDate(txnDate)

    If (IsMonthClosed(monthKey) And Not allowInClosedMonth) Then
        Err.Raise vbObjectError + 513, "AddLedgerEntry", "Month is closed: " & monthKey
    End If

    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    Dim lr As ListRow: Set lr = lo.ListRows.Add

    Dim txnId As String: txnId = NextTxnId(monthKey)
    Dim fy As Long: fy = FiscalYearForMonthKey(monthKey)

    lr.Range.Cells(1, lo.ListColumns("TxnID").Index).value = txnId
    lr.Range.Cells(1, lo.ListColumns("Date").Index).value = txnDate
    lr.Range.Cells(1, lo.ListColumns("MonthKey").Index).value = monthKey
    lr.Range.Cells(1, lo.ListColumns("FiscalYear").Index).value = fy

    lr.Range.Cells(1, lo.ListColumns("TxnType").Index).value = txnType
    lr.Range.Cells(1, lo.ListColumns("TxnDetail").Index).value = txnDetail

    lr.Range.Cells(1, lo.ListColumns("Category").Index).value = category
    lr.Range.Cells(1, lo.ListColumns("Event").Index).value = eventName
    lr.Range.Cells(1, lo.ListColumns("Charity").Index).value = charityName

    lr.Range.Cells(1, lo.ListColumns("Gross").Index).value = gross
    lr.Range.Cells(1, lo.ListColumns("Fees").Index).value = fees
    lr.Range.Cells(1, lo.ListColumns("Net").Index).value = gross - fees

    lr.Range.Cells(1, lo.ListColumns("PaymentMethod").Index).value = paymentMethod
    lr.Range.Cells(1, lo.ListColumns("PayeeOrSource").Index).value = payeeOrSource
    lr.Range.Cells(1, lo.ListColumns("Memo").Index).value = memo

    lr.Range.Cells(1, lo.ListColumns("ReceiptRequired").Index).value = receiptRequired
    lr.Range.Cells(1, lo.ListColumns("ReceiptStatus").Index).value = IIf(receiptRequired, "Missing", "NotRequired")
    lr.Range.Cells(1, lo.ListColumns("ReceiptLink").Index).value = ""

    lr.Range.Cells(1, lo.ListColumns("ApprovedBy").Index).value = GetConfigValue(CFG_APPROVER_NAME, Application.UserName)
    lr.Range.Cells(1, lo.ListColumns("ClosedFlag").Index).value = False

    lr.Range.Cells(1, lo.ListColumns("CreatedAt").Index).value = Now
    lr.Range.Cells(1, lo.ListColumns("UpdatedAt").Index).value = Now

    AuditLog "Create", txnId, txnType & " / " & monthKey & " / " & Format$(gross - fees, "0.00")

    AddLedgerEntry = txnId
End Function

Public Sub UpdateLedgerFields(ByVal txnId As String, ByVal category As String, ByVal eventName As String, ByVal charityName As String, _
                             ByVal receiptRequired As Boolean)
    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    Dim rowIdx As Long: rowIdx = FindLedgerRowIndex(txnId)
    If rowIdx = 0 Then Err.Raise vbObjectError + 514, "UpdateLedgerFields", "TxnID not found: " & txnId

    Dim monthKey As String
    monthKey = CStr(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("MonthKey").Index).value)

    If IsMonthClosed(monthKey) Then
        Err.Raise vbObjectError + 515, "UpdateLedgerFields", "Month is closed: " & monthKey
    End If

    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("Category").Index).value = category
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("Event").Index).value = eventName
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("Charity").Index).value = charityName

    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ReceiptRequired").Index).value = receiptRequired

    Dim statusCol As Long: statusCol = lo.ListColumns("ReceiptStatus").Index
    Dim linkCol As Long: linkCol = lo.ListColumns("ReceiptLink").Index

    Dim curLink As String: curLink = NzStr(lo.DataBodyRange.Cells(rowIdx, linkCol).value, "")
    Dim curStatus As String: curStatus = NzStr(lo.DataBodyRange.Cells(rowIdx, statusCol).value, "")

    If receiptRequired Then
        If Len(curLink) = 0 Then
            If curStatus <> "Waived" Then lo.DataBodyRange.Cells(rowIdx, statusCol).value = "Missing"
        Else
            lo.DataBodyRange.Cells(rowIdx, statusCol).value = "Linked"
        End If
    Else
        lo.DataBodyRange.Cells(rowIdx, statusCol).value = "NotRequired"
    End If

    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("UpdatedAt").Index).value = Now
    AuditLog "Edit", txnId, "Category/Event/Charity/ReceiptRequired"
End Sub

Private Function FindLedgerRowIndex(ByVal txnId As String) As Long
    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    FindLedgerRowIndex = 0
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Range
    For Each r In lo.ListColumns("TxnID").DataBodyRange.Cells
        If CStr(r.value) = txnId Then
            FindLedgerRowIndex = r.Row - lo.DataBodyRange.Row + 1
            Exit Function
        End If
    Next r
End Function

'========================
' Receipts
'========================

Public Sub AttachReceiptToTxn(ByVal txnId As String)
    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    Dim rowIdx As Long: rowIdx = FindLedgerRowIndex(txnId)
    If rowIdx = 0 Then Err.Raise vbObjectError + 516, "AttachReceiptToTxn", "TxnID not found: " & txnId

    Dim monthKey As String: monthKey = CStr(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("MonthKey").Index).value)
    If IsMonthClosed(monthKey) Then Err.Raise vbObjectError + 517, "AttachReceiptToTxn", "Month is closed: " & monthKey

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = False
    fd.Title = "Select receipt file"
    fd.Filters.Clear
    fd.Filters.Add "Images/PDF", "*.pdf;*.jpg;*.jpeg;*.png;*.heic", 1

    If fd.Show <> -1 Then Exit Sub ' canceled

    Dim src As String: src = fd.SelectedItems(1)
    Dim destFolder As String: destFolder = ReceiptFolderForMonth(monthKey)
    EnsureFolderPath destFolder

    Dim d As Date: d = CDate(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("Date").Index).value)
    Dim payee As String: payee = SanitizeFilePart(NzStr(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("PayeeOrSource").Index).value, "Receipt"))
    Dim ext As String: ext = LCase$(Mid$(src, InStrRev(src, ".")))

    Dim destName As String
    destName = txnId & "__" & Format$(d, "yyyy-mm-dd") & "__" & payee & ext

    Dim dest As String: dest = destFolder & destName
    FileCopy src, dest

    Dim linkCol As Long: linkCol = lo.ListColumns("ReceiptLink").Index
    lo.DataBodyRange.Cells(rowIdx, linkCol).value = dest

    ' set hyperlink with display text
    Dim cell As Range: Set cell = lo.DataBodyRange.Cells(rowIdx, linkCol)
    On Error Resume Next
    cell.Hyperlinks.Delete
    On Error GoTo 0
    ThisWorkbook.Worksheets(SH_LEDGER).Hyperlinks.Add Anchor:=cell, Address:=dest, TextToDisplay:="Receipt"

    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ReceiptStatus").Index).value = "Linked"
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("UpdatedAt").Index).value = Now

    AuditLog "AttachReceipt", txnId, dest
End Sub

Public Sub WaiveReceipt(ByVal txnId As String, ByVal reason As String)
    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    Dim rowIdx As Long: rowIdx = FindLedgerRowIndex(txnId)
    If rowIdx = 0 Then Err.Raise vbObjectError + 518, "WaiveReceipt", "TxnID not found: " & txnId

    Dim monthKey As String: monthKey = CStr(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("MonthKey").Index).value)
    If IsMonthClosed(monthKey) Then Err.Raise vbObjectError + 519, "WaiveReceipt", "Month is closed: " & monthKey

    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ReceiptStatus").Index).value = "Waived"
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ReceiptWaivedReason").Index).value = reason
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ReceiptWaivedBy").Index).value = GetConfigValue(CFG_APPROVER_NAME, Application.UserName)
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ReceiptWaivedAt").Index).value = Now
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("UpdatedAt").Index).value = Now

    AuditLog "WaiveReceipt", txnId, reason
End Sub

Private Function ReceiptFolderForMonth(ByVal monthKey As String) As String
    Dim baseRel As String: baseRel = GetConfigValue(CFG_RECEIPTS_FOLDER, ".\Receipts\")
    Dim baseAbs As String: baseAbs = ResolveWorkbookRelativePath(baseRel)
    Dim fy As Long: fy = FiscalYearForMonthKey(monthKey)

    ReceiptFolderForMonth = EnsureTrailingSlash(baseAbs) & "FY" & CStr(fy) & "\" & monthKey & "\"
End Function

Private Function BoardPacketsFolderAbs() As String
    BoardPacketsFolderAbs = ResolveWorkbookRelativePath(GetConfigValue(CFG_BOARDPACKETS_FOLDER, ".\BoardPackets\"))
    BoardPacketsFolderAbs = EnsureTrailingSlash(BoardPacketsFolderAbs)
End Function

Private Function ResolveWorkbookRelativePath(ByVal rel As String) As String
    Dim p As String: p = ThisWorkbook.path
    If Len(p) = 0 Then p = CurDir$
    ResolveWorkbookRelativePath = EnsureTrailingSlash(p) & Replace(rel, ".\", "")
End Function

Private Function EnsureTrailingSlash(ByVal p As String) As String
    If Right$(p, 1) = "\" Then
        EnsureTrailingSlash = p
    Else
        EnsureTrailingSlash = p & "\"
    End If
End Function

Private Sub EnsureFolderPath(ByVal folderPath As String)
    Dim parts() As String: parts = Split(folderPath, "\")
    Dim i As Long, cur As String

    If InStr(folderPath, ":\") > 0 Then
        cur = parts(0) & "\"
        i = 1
    Else
        cur = ""
        i = 0
    End If

    For i = i To UBound(parts)
        If Len(parts(i)) > 0 Then
            cur = cur & parts(i) & "\"
            If Dir(cur, vbDirectory) = "" Then MkDir cur
        End If
    Next i
End Sub

Private Function SanitizeFilePart(ByVal s As String) As String
    Dim bad As Variant: bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "_")
    Next i
    s = Trim$(s)
    If Len(s) = 0 Then s = "Receipt"
    If Len(s) > 60 Then s = Left$(s, 60)
    SanitizeFilePart = s
End Function

'========================
' Month status / reconcile / close
'========================

Public Function IsMonthClosed(ByVal monthKey As String) As Boolean
    Dim lo As ListObject: Set lo = GetTable(SH_MONTHSTATUS, T_MONTHSTATUS)
    Dim rowIdx As Long: rowIdx = FindMonthStatusRow(monthKey)
    If rowIdx = 0 Then
        IsMonthClosed = False
    Else
        IsMonthClosed = CBool(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ClosedFlag").Index).value)
    End If
End Function

Private Function FindMonthStatusRow(ByVal monthKey As String) As Long
    Dim lo As ListObject: Set lo = GetTable(SH_MONTHSTATUS, T_MONTHSTATUS)
    FindMonthStatusRow = 0
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Range
    For Each r In lo.ListColumns("MonthKey").DataBodyRange.Cells
        If CStr(r.value) = monthKey Then
            FindMonthStatusRow = r.Row - lo.DataBodyRange.Row + 1
            Exit Function
        End If
    Next r
End Function

Private Sub EnsureMonthStatusRow(ByVal monthKey As String)
    Dim lo As ListObject: Set lo = GetTable(SH_MONTHSTATUS, T_MONTHSTATUS)
    If FindMonthStatusRow(monthKey) <> 0 Then Exit Sub

    Dim fy As Long: fy = FiscalYearForMonthKey(monthKey)
    Dim lr As ListRow: Set lr = lo.ListRows.Add
    lr.Range.Cells(1, lo.ListColumns("MonthKey").Index).value = monthKey
    lr.Range.Cells(1, lo.ListColumns("FiscalYear").Index).value = fy
    lr.Range.Cells(1, lo.ListColumns("BeginningBalance").Index).value = 0#
    lr.Range.Cells(1, lo.ListColumns("EndingBalance").Index).value = 0#
    lr.Range.Cells(1, lo.ListColumns("LedgerDeposits").Index).value = 0#
    lr.Range.Cells(1, lo.ListColumns("LedgerWithdrawals").Index).value = 0#
    lr.Range.Cells(1, lo.ListColumns("ExpectedEnding").Index).value = 0#
    lr.Range.Cells(1, lo.ListColumns("ReconDifference").Index).value = 0#
    lr.Range.Cells(1, lo.ListColumns("ReconStatus").Index).value = "NotRun"
    lr.Range.Cells(1, lo.ListColumns("LastReconAt").Index).value = ""
    lr.Range.Cells(1, lo.ListColumns("ClosedFlag").Index).value = False
    lr.Range.Cells(1, lo.ListColumns("ClosedAt").Index).value = ""
End Sub

Public Sub ComputeMonthLedgerTotals(ByVal monthKey As String, ByRef deposits As Double, ByRef withdrawals As Double)
    deposits = 0#
    withdrawals = 0#

    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim i As Long
    Dim mkCol As Long: mkCol = lo.ListColumns("MonthKey").Index
    Dim typeCol As Long: typeCol = lo.ListColumns("TxnType").Index
    Dim netCol As Long: netCol = lo.ListColumns("Net").Index

    For i = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(i, mkCol).value) = monthKey Then
            Dim t As String: t = CStr(lo.DataBodyRange.Cells(i, typeCol).value)
            Dim n As Double: n = CDbl(lo.DataBodyRange.Cells(i, netCol).value)

            If LCase$(t) = "income" Then
                deposits = deposits + n
            ElseIf LCase$(t) = "expense" Or LCase$(t) = "reimbursement" Then
                withdrawals = withdrawals + Abs(n)
            ElseIf LCase$(t) = "adjustment" Then
                If n >= 0 Then deposits = deposits + n Else withdrawals = withdrawals + Abs(n)
            End If
        End If
    Next i
End Sub

Public Sub SaveReconciliation(ByVal monthKey As String, ByVal beginningBal As Double, ByVal endingBal As Double)
    EnsureMonthStatusRow monthKey

    Dim deposits As Double, withdrawals As Double
    ComputeMonthLedgerTotals monthKey, deposits, withdrawals

    Dim expected As Double: expected = beginningBal + deposits - withdrawals
    Dim diff As Double: diff = expected - endingBal

    Dim lo As ListObject: Set lo = GetTable(SH_MONTHSTATUS, T_MONTHSTATUS)
    Dim rowIdx As Long: rowIdx = FindMonthStatusRow(monthKey)

    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("BeginningBalance").Index).value = beginningBal
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("EndingBalance").Index).value = endingBal
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("LedgerDeposits").Index).value = deposits
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("LedgerWithdrawals").Index).value = withdrawals
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ExpectedEnding").Index).value = expected
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ReconDifference").Index).value = diff
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ReconStatus").Index).value = IIf(Abs(diff) < 0.005, "OK", "OutOfBalance")
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("LastReconAt").Index).value = Now

    AuditLog "Reconcile", "", monthKey & " diff=" & Format$(diff, "0.00")
End Sub

Public Function GateCheckMonth(ByVal monthKey As String, ByRef missingCategories As Long, ByRef missingReceipts As Long, ByRef missingReceiptsAmt As Double, ByRef reconOk As Boolean) As String
    missingCategories = 0
    missingReceipts = 0
    missingReceiptsAmt = 0#
    reconOk = False

    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    Dim mkCol As Long: mkCol = lo.ListColumns("MonthKey").Index
    Dim catCol As Long: catCol = lo.ListColumns("Category").Index
    Dim rrCol As Long: rrCol = lo.ListColumns("ReceiptRequired").Index
    Dim rsCol As Long: rsCol = lo.ListColumns("ReceiptStatus").Index
    Dim netCol As Long: netCol = lo.ListColumns("Net").Index

    If Not lo.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = 1 To lo.ListRows.count
            If CStr(lo.DataBodyRange.Cells(i, mkCol).value) = monthKey Then
                If Len(Trim$(CStr(lo.DataBodyRange.Cells(i, catCol).value))) = 0 Then
                    missingCategories = missingCategories + 1
                End If

                Dim rr As Boolean: rr = CBool(lo.DataBodyRange.Cells(i, rrCol).value)
                Dim rs As String: rs = CStr(lo.DataBodyRange.Cells(i, rsCol).value)

                If rr Then
                    If rs <> "Linked" And rs <> "Waived" Then
                        missingReceipts = missingReceipts + 1
                        missingReceiptsAmt = missingReceiptsAmt + Abs(CDbl(lo.DataBodyRange.Cells(i, netCol).value))
                    End If
                End If
            End If
        Next i
    End If

    reconOk = IsReconOk(monthKey)

    Dim msg As String
    msg = "Month " & monthKey & " gates:" & vbCrLf & _
          "Uncategorized: " & CStr(missingCategories) & vbCrLf & _
          "Missing receipts: " & CStr(missingReceipts) & " ($" & Format$(missingReceiptsAmt, "0.00") & ")" & vbCrLf & _
          "Reconciled: " & IIf(reconOk, "YES", "NO")
    GateCheckMonth = msg
End Function

Public Function IsReconOk(ByVal monthKey As String) As Boolean
    Dim lo As ListObject: Set lo = GetTable(SH_MONTHSTATUS, T_MONTHSTATUS)
    Dim rowIdx As Long: rowIdx = FindMonthStatusRow(monthKey)
    If rowIdx = 0 Then
        IsReconOk = False
        Exit Function
    End If

    Dim status As String: status = CStr(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ReconStatus").Index).value)
    IsReconOk = (status = "OK")
End Function

Public Sub CloseMonth(ByVal monthKey As String)
    Dim mc As Long, mr As Long
    Dim mAmt As Double, reconOk As Boolean
    Dim gateMsg As String
    gateMsg = GateCheckMonth(monthKey, mc, mr, mAmt, reconOk)

    If mc <> 0 Or mr <> 0 Or (Not reconOk) Then
        Err.Raise vbObjectError + 520, "CloseMonth", "Close blocked." & vbCrLf & gateMsg
    End If

    EnsureMonthStatusRow monthKey

    Dim ms As ListObject: Set ms = GetTable(SH_MONTHSTATUS, T_MONTHSTATUS)
    Dim rowIdx As Long: rowIdx = FindMonthStatusRow(monthKey)
    ms.DataBodyRange.Cells(rowIdx, ms.ListColumns("ClosedFlag").Index).value = True
    ms.DataBodyRange.Cells(rowIdx, ms.ListColumns("ClosedAt").Index).value = Now

    ' mark ledger rows closed
    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    If Not lo.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = 1 To lo.ListRows.count
            If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MonthKey").Index).value) = monthKey Then
                lo.DataBodyRange.Cells(i, lo.ListColumns("ClosedFlag").Index).value = True
            End If
        Next i
    End If

    ProtectDataSheets
    AuditLog "CloseMonth", "", monthKey
End Sub

Private Sub ProtectDataSheets()
    Dim locksEnabled As String: locksEnabled = UCase$(GetConfigValue(CFG_LOCKS_ENABLED, "TRUE"))
    If locksEnabled <> "TRUE" Then Exit Sub

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> SH_HOME Then
            On Error Resume Next
            ws.Protect Password:="", UserInterfaceOnly:=True, AllowFiltering:=True
            On Error GoTo 0
        End If
    Next ws
End Sub

'========================
' Reporting (Monthly packet)
'========================

Public Sub GenerateMonthlyPacket(ByVal monthKey As String)
    Dim prevSU As Boolean: prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim ws As Worksheet: Set ws = GetSheet(SH_REPORT)
    Dim oldVis As XlSheetVisibility: oldVis = ws.Visible
    ws.Visible = xlSheetVisible

    EnsureReportSheetLayout

    Dim fy As Long: fy = FiscalYearForMonthKey(monthKey)
    ws.Range("B3").value = monthKey
    ws.Range("B4").value = fy

    ' balances from month status (if present)
    Dim beginCash As Double: beginCash = GetMonthStatusValue(monthKey, "BeginningBalance")
    Dim endCash As Double: endCash = GetMonthStatusValue(monthKey, "EndingBalance")

    ' month totals
    Dim inc As Double: inc = SumLedgerNet(monthKey, "Income", "(All)", "(All)", "(All)")
    Dim exp As Double: exp = SumLedgerNet(monthKey, "Expense", "(All)", "(All)", "(All)") + SumLedgerNet(monthKey, "Reimbursement", "(All)", "(All)", "(All)")
    Dim netChg As Double: netChg = inc - exp

    ws.Range("B7").value = beginCash
    ws.Range("B8").value = inc
    ws.Range("B9").value = exp
    ws.Range("B10").value = netChg
    ws.Range("B11").value = endCash

    ' controls snapshot
    ws.Range("E7").value = IIf(IsReconOk(monthKey), "YES", "NO")
    ws.Range("E8").value = IIf(IsMonthClosed(monthKey), "YES", "NO")

    Dim mc As Long, mr As Long, mAmt As Double, reconOk As Boolean
    Call GateCheckMonth(monthKey, mc, mr, mAmt, reconOk)
    ws.Range("E9").value = CStr(mr) & " / $" & Format$(mAmt, "0.00")
    ws.Range("E10").value = CStr(mc)

    ' Budget vs Actual
    Dim categories As Variant
    categories = Array("Administrative", "Programs", "Fundraising", "Marketing", "Travel", "Services", "Misc")

    Dim r As Long: r = 15
    Dim i As Long
    For i = LBound(categories) To UBound(categories)
        Dim cat As String: cat = CStr(categories(i))
        ws.Cells(r, 1).value = cat

        Dim b As Double: b = GetBudget(monthKey, cat)
        Dim a As Double: a = SumLedgerNet(monthKey, "(All)", "(All)", "(All)", cat) ' actual expenses by category (net absolute)
        ws.Cells(r, 2).value = b
        ws.Cells(r, 3).value = a
        ws.Cells(r, 4).value = b - a
        If b = 0 Then
            ws.Cells(r, 5).value = ""
        Else
            ws.Cells(r, 5).value = (b - a) / b
        End If
        r = r + 1
    Next i

    ' charity month + ytd
    Dim charityRaisedM As Double: charityRaisedM = SumCharity(monthKey, "Raised")
    Dim charityPaidM As Double: charityPaidM = SumCharity(monthKey, "Paid")
    Dim CharityHeldYTD As Double: CharityHeldYTD = SumCharityYTD(monthKey, "Raised") - SumCharityYTD(monthKey, "Paid")

    ws.Range("H14").value = charityRaisedM
    ws.Range("H15").value = charityPaidM
    ws.Range("H16").value = CharityHeldYTD

    ' event rollup
    WriteEventRollup ws, monthKey

    ' ytd section
    Dim ytdInc As Double: ytdInc = SumLedgerNetYTD(monthKey, "Income")
    Dim ytdExp As Double: ytdExp = SumLedgerNetYTD(monthKey, "Expense") + SumLedgerNetYTD(monthKey, "Reimbursement")
    ws.Range("B36").value = ytdInc
    ws.Range("B37").value = ytdExp
    ws.Range("B38").value = ytdInc - ytdExp

    Dim ytdBudgetVar As Double: ytdBudgetVar = BudgetVarYTD(monthKey)
    ws.Range("B39").value = ytdBudgetVar

    ' format
    ws.Columns("A:H").AutoFit
    ws.Range("B7:B11,H14:H16,B36:B39").NumberFormat = "$#,##0.00"
    ws.Range("B15:D21").NumberFormat = "$#,##0.00"
    ws.Range("E15:E21").NumberFormat = "0.0%"

    ' export pdf
    Dim outFolder As String: outFolder = BoardPacketsFolderAbs()
    EnsureFolderPath outFolder
    Dim pdfPath As String: pdfPath = outFolder & "TCPP_BoardPacket_" & monthKey & ".pdf"

    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    ws.Visible = oldVis
    AuditLog "GenerateMonthlyPacket", "", pdfPath

    Application.ScreenUpdating = prevSU
End Sub

Private Function GetMonthStatusValue(ByVal monthKey As String, ByVal colName As String) As Double
    Dim lo As ListObject: Set lo = GetTable(SH_MONTHSTATUS, T_MONTHSTATUS)
    Dim rowIdx As Long: rowIdx = FindMonthStatusRow(monthKey)
    If rowIdx = 0 Then
        GetMonthStatusValue = 0#
    Else
        GetMonthStatusValue = NzDbl(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns(colName).Index).value, 0#)
    End If
End Function

Private Function SumLedgerNet(ByVal monthKey As String, ByVal txnType As String, ByVal eventFilter As String, ByVal charityFilter As String, ByVal categoryFilter As String) As Double
    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    SumLedgerNet = 0#
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim mkCol As Long: mkCol = lo.ListColumns("MonthKey").Index
    Dim typeCol As Long: typeCol = lo.ListColumns("TxnType").Index
    Dim evtCol As Long: evtCol = lo.ListColumns("Event").Index
    Dim chCol As Long: chCol = lo.ListColumns("Charity").Index
    Dim catCol As Long: catCol = lo.ListColumns("Category").Index
    Dim netCol As Long: netCol = lo.ListColumns("Net").Index

    Dim i As Long
    For i = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(i, mkCol).value) = monthKey Then
            Dim t As String: t = CStr(lo.DataBodyRange.Cells(i, typeCol).value)

            If txnType <> "(All)" Then
                If LCase$(t) <> LCase$(txnType) Then GoTo ContinueRow
            End If

            Dim ev As String: ev = NzStr(lo.DataBodyRange.Cells(i, evtCol).value, "")
            Dim ch As String: ch = NzStr(lo.DataBodyRange.Cells(i, chCol).value, "")
            Dim cat As String: cat = NzStr(lo.DataBodyRange.Cells(i, catCol).value, "")

            If eventFilter <> "(All)" And eventFilter <> ev Then GoTo ContinueRow
            If charityFilter <> "(All)" And charityFilter <> ch Then GoTo ContinueRow

            If categoryFilter <> "(All)" Then
                ' category filter represents EXPENSE categorization; treat it as absolute outflow sum for Expense/Reimbursement
                If cat <> categoryFilter Then GoTo ContinueRow
                If LCase$(t) = "expense" Or LCase$(t) = "reimbursement" Then
                    SumLedgerNet = SumLedgerNet + Abs(CDbl(lo.DataBodyRange.Cells(i, netCol).value))
                End If
            Else
                ' normal net sum by type
                If LCase$(t) = "income" Then
                    SumLedgerNet = SumLedgerNet + CDbl(lo.DataBodyRange.Cells(i, netCol).value)
                ElseIf LCase$(t) = "expense" Or LCase$(t) = "reimbursement" Then
                    SumLedgerNet = SumLedgerNet + Abs(CDbl(lo.DataBodyRange.Cells(i, netCol).value))
                ElseIf LCase$(t) = "adjustment" Then
                    ' treat adjustments as net absolute if used for expense categories; otherwise ignore here
                End If
            End If
        End If
ContinueRow:
    Next i
End Function

Private Function FiscalYearStartMonthKey(ByVal anyMonthKey As String) As String
    Dim fy As Long: fy = FiscalYearForMonthKey(anyMonthKey)
    Dim startM As Long: startM = CLng(GetConfigValue(CFG_FISCAL_START_MONTH, "6"))
    Dim startYear As Long: startYear = fy - 1
    FiscalYearStartMonthKey = Format$(DateSerial(startYear, startM, 1), "yyyy-mm")
End Function

Private Function MonthKeyLessOrEqual(ByVal a As String, ByVal b As String) As Boolean
    MonthKeyLessOrEqual = (CLng(Replace(a, "-", "")) <= CLng(Replace(b, "-", "")))
End Function

Private Function MonthKeyAdd(ByVal monthKey As String, ByVal months As Long) As String
    Dim y As Long, m As Long
    y = CLng(Left$(monthKey, 4))
    m = CLng(Right$(monthKey, 2))
    MonthKeyAdd = Format$(DateAdd("m", months, DateSerial(y, m, 1)), "yyyy-mm")
End Function

Private Function SumLedgerNetYTD(ByVal monthKey As String, ByVal txnType As String) As Double
    Dim startKey As String: startKey = FiscalYearStartMonthKey(monthKey)
    Dim cur As String: cur = startKey
    SumLedgerNetYTD = 0#
    Do While MonthKeyLessOrEqual(cur, monthKey)
        If LCase$(txnType) = "income" Then
            SumLedgerNetYTD = SumLedgerNetYTD + SumLedgerNet(cur, "Income", "(All)", "(All)", "(All)")
        ElseIf LCase$(txnType) = "expense" Then
            SumLedgerNetYTD = SumLedgerNetYTD + SumLedgerNet(cur, "Expense", "(All)", "(All)", "(All)")
        ElseIf LCase$(txnType) = "reimbursement" Then
            SumLedgerNetYTD = SumLedgerNetYTD + SumLedgerNet(cur, "Reimbursement", "(All)", "(All)", "(All)")
        End If
        cur = MonthKeyAdd(cur, 1)
    Loop
End Function

Private Function SumCharity(ByVal monthKey As String, ByVal mode As String) As Double
    ' Raised = Income rows with Charity set
    ' Paid = Expense/Reimbursement rows with Charity set (absolute)
    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    SumCharity = 0#
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim mkCol As Long: mkCol = lo.ListColumns("MonthKey").Index
    Dim typeCol As Long: typeCol = lo.ListColumns("TxnType").Index
    Dim chCol As Long: chCol = lo.ListColumns("Charity").Index
    Dim netCol As Long: netCol = lo.ListColumns("Net").Index

    Dim i As Long
    For i = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(i, mkCol).value) = monthKey Then
            Dim ch As String: ch = NzStr(lo.DataBodyRange.Cells(i, chCol).value, "")
            If Len(ch) = 0 Then GoTo ContinueRow

            Dim t As String: t = LCase$(CStr(lo.DataBodyRange.Cells(i, typeCol).value))
            Dim n As Double: n = CDbl(lo.DataBodyRange.Cells(i, netCol).value)

            If mode = "Raised" Then
                If t = "income" Then SumCharity = SumCharity + n
            ElseIf mode = "Paid" Then
                If t = "expense" Or t = "reimbursement" Then SumCharity = SumCharity + Abs(n)
            End If
        End If
ContinueRow:
    Next i
End Function

Private Function SumCharityYTD(ByVal monthKey As String, ByVal mode As String) As Double
    Dim startKey As String: startKey = FiscalYearStartMonthKey(monthKey)
    Dim cur As String: cur = startKey
    SumCharityYTD = 0#
    Do While MonthKeyLessOrEqual(cur, monthKey)
        SumCharityYTD = SumCharityYTD + SumCharity(cur, mode)
        cur = MonthKeyAdd(cur, 1)
    Loop
End Function

Private Sub WriteEventRollup(ByVal ws As Worksheet, ByVal monthKey As String)
    ' clear area
    ws.Range("A26:D34").ClearContents

    Dim evLo As ListObject: Set evLo = GetTable(SH_LOOKUPS, T_EVENTS)
    If evLo.DataBodyRange Is Nothing Then Exit Sub

    Dim r As Long: r = 26
    Dim cell As Range
    For Each cell In evLo.ListColumns(1).DataBodyRange.Cells
        Dim ev As String: ev = CStr(cell.value)

        Dim inc As Double: inc = SumLedgerNet(monthKey, "Income", ev, "(All)", "(All)")
        Dim exp As Double: exp = SumLedgerNet(monthKey, "Expense", ev, "(All)", "(All)") + SumLedgerNet(monthKey, "Reimbursement", ev, "(All)", "(All)")
        Dim net As Double: net = inc - exp

        If Abs(inc) > 0.005 Or Abs(exp) > 0.005 Then
            ws.Cells(r, 1).value = ev
            ws.Cells(r, 2).value = inc
            ws.Cells(r, 3).value = exp
            ws.Cells(r, 4).value = net
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).NumberFormat = "$#,##0.00"
            r = r + 1
            If r > 34 Then Exit Sub
        End If
    Next cell
End Sub

'========================
' Budget
'========================

Public Function GetBudget(ByVal monthKey As String, ByVal category As String) As Double
    Dim lo As ListObject: Set lo = GetTable(SH_BUDGET, T_BUDGET)
    GetBudget = 0#
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim mkCol As Long: mkCol = lo.ListColumns("MonthKey").Index
    Dim catCol As Long: catCol = lo.ListColumns("Category").Index
    Dim amtCol As Long: amtCol = lo.ListColumns("BudgetAmount").Index

    Dim i As Long
    For i = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(i, mkCol).value) = monthKey And CStr(lo.DataBodyRange.Cells(i, catCol).value) = category Then
            GetBudget = NzDbl(lo.DataBodyRange.Cells(i, amtCol).value, 0#)
            Exit Function
        End If
    Next i
End Function

Public Sub SetBudget(ByVal monthKey As String, ByVal category As String, ByVal amount As Double)
    Dim lo As ListObject: Set lo = GetTable(SH_BUDGET, T_BUDGET)

    Dim fy As Long: fy = FiscalYearForMonthKey(monthKey)

    If lo.DataBodyRange Is Nothing Then
        Dim lr0 As ListRow: Set lr0 = lo.ListRows.Add
        lr0.Range.Cells(1, lo.ListColumns("MonthKey").Index).value = monthKey
        lr0.Range.Cells(1, lo.ListColumns("FiscalYear").Index).value = fy
        lr0.Range.Cells(1, lo.ListColumns("Category").Index).value = category
        lr0.Range.Cells(1, lo.ListColumns("BudgetAmount").Index).value = amount
        Exit Sub
    End If

    Dim mkCol As Long: mkCol = lo.ListColumns("MonthKey").Index
    Dim catCol As Long: catCol = lo.ListColumns("Category").Index
    Dim amtCol As Long: amtCol = lo.ListColumns("BudgetAmount").Index
    Dim fyCol As Long: fyCol = lo.ListColumns("FiscalYear").Index

    Dim i As Long
    For i = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(i, mkCol).value) = monthKey And CStr(lo.DataBodyRange.Cells(i, catCol).value) = category Then
            lo.DataBodyRange.Cells(i, amtCol).value = amount
            lo.DataBodyRange.Cells(i, fyCol).value = fy
            Exit Sub
        End If
    Next i

    Dim lr As ListRow: Set lr = lo.ListRows.Add
    lr.Range.Cells(1, mkCol).value = monthKey
    lr.Range.Cells(1, fyCol).value = fy
    lr.Range.Cells(1, catCol).value = category
    lr.Range.Cells(1, amtCol).value = amount
End Sub

Private Function BudgetVarYTD(ByVal monthKey As String) As Double
    Dim startKey As String: startKey = FiscalYearStartMonthKey(monthKey)
    Dim cur As String: cur = startKey

    Dim cats As Variant
    cats = Array("Administrative", "Programs", "Fundraising", "Marketing", "Travel", "Services", "Misc")

    Dim varSum As Double: varSum = 0#

    Do While MonthKeyLessOrEqual(cur, monthKey)
        Dim i As Long
        For i = LBound(cats) To UBound(cats)
            Dim cat As String: cat = CStr(cats(i))
            Dim b As Double: b = GetBudget(cur, cat)
            Dim a As Double: a = SumLedgerNet(cur, "(All)", "(All)", "(All)", cat)
            varSum = varSum + (b - a)
        Next i
        cur = MonthKeyAdd(cur, 1)
    Loop

    BudgetVarYTD = varSum
End Function

'========================
' Dashboard calculations
'========================

Public Sub GetExceptionCounts(ByVal monthKey As String, ByRef cntUncategorized As Long, ByRef cntMissingReceipts As Long, ByRef amtMissingReceipts As Double)
    Dim mc As Long, mr As Long
    Dim mAmt As Double, reconOk As Boolean
    Call GateCheckMonth(monthKey, mc, mr, mAmt, reconOk)
    cntUncategorized = mc
    cntMissingReceipts = mr
    amtMissingReceipts = mAmt
End Sub

Public Function CharityHeldYTD(ByVal monthKey As String) As Double
    CharityHeldYTD = SumCharityYTD(monthKey, "Raised") - SumCharityYTD(monthKey, "Paid")
End Function

Public Function SelfTestReport() As String
    Dim ok As Boolean: ok = True
    Dim msg As String: msg = ""

    msg = msg & "Sheets:" & vbCrLf
    msg = msg & "- " & SH_HOME & ": " & SheetExists(SH_HOME) & vbCrLf
    msg = msg & "- " & SH_LEDGER & ": " & SheetExists(SH_LEDGER) & vbCrLf
    msg = msg & "- " & SH_LOOKUPS & ": " & SheetExists(SH_LOOKUPS) & vbCrLf
    msg = msg & "- " & SH_BUDGET & ": " & SheetExists(SH_BUDGET) & vbCrLf
    msg = msg & "- " & SH_MONTHSTATUS & ": " & SheetExists(SH_MONTHSTATUS) & vbCrLf
    msg = msg & "- " & SH_AUDIT & ": " & SheetExists(SH_AUDIT) & vbCrLf
    msg = msg & "- " & SH_REPORT & ": " & SheetExists(SH_REPORT) & vbCrLf & vbCrLf

    msg = msg & "Tables:" & vbCrLf
    msg = msg & "- " & T_LEDGER & ": " & TableExists(SH_LEDGER, T_LEDGER) & vbCrLf
    msg = msg & "- " & T_COA & ": " & TableExists(SH_LOOKUPS, T_COA) & vbCrLf
    msg = msg & "- " & T_EVENTS & ": " & TableExists(SH_LOOKUPS, T_EVENTS) & vbCrLf
    msg = msg & "- " & T_CHARITIES & ": " & TableExists(SH_LOOKUPS, T_CHARITIES) & vbCrLf
    msg = msg & "- " & T_PAYMETHOD & ": " & TableExists(SH_LOOKUPS, T_PAYMETHOD) & vbCrLf
    msg = msg & "- " & T_CONFIG & ": " & TableExists(SH_LOOKUPS, T_CONFIG) & vbCrLf
    msg = msg & "- " & T_BUDGET & ": " & TableExists(SH_BUDGET, T_BUDGET) & vbCrLf
    msg = msg & "- " & T_MONTHSTATUS & ": " & TableExists(SH_MONTHSTATUS, T_MONTHSTATUS) & vbCrLf
    msg = msg & "- " & T_AUDIT & ": " & TableExists(SH_AUDIT, T_AUDIT) & vbCrLf

    SelfTestReport = msg
End Function

Private Function SheetExists(ByVal name As String) As String
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(name)
    SheetExists = IIf(ws Is Nothing, "MISSING", "OK")
End Function

Private Function TableExists(ByVal sheetName As String, ByVal tableName As String) As String
    On Error Resume Next
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    TableExists = IIf(lo Is Nothing, "MISSING", "OK")
End Function



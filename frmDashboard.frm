VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDashboard 
   Caption         =   "UserForm1"
   ClientHeight    =   7160
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12045
   OleObjectBlob   =   "frmDashboard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    On Error GoTo EH

    modTCPPv2.InitializeTool False
    modTCPPv2.ApplyTheme Me

    LoadMonthList
    LoadEventList
    LoadCharityList

    cboMonth.value = modTCPPv2.gMonthKey
    cboEvent.value = modTCPPv2.gEventFilter
    cboCharity.value = modTCPPv2.gCharityFilter

    lstExceptions.ColumnCount = 2
    lstExceptions.ColumnWidths = "220;60"

    RefreshDashboard
    Exit Sub
EH:
    modTCPPv2.HandleError "frmDashboard.Initialize", Err, ""
End Sub

Private Sub cmdRefreshDashboard_Click()
    RefreshDashboard
End Sub

Private Sub cboMonth_Change()
    modTCPPv2.gMonthKey = NzCombo(cboMonth.value, Format(Date, "yyyy-mm"))
    RefreshDashboard
End Sub

Private Sub cboEvent_Change()
    modTCPPv2.gEventFilter = NzCombo(cboEvent.value, "(All)")
    RefreshDashboard
End Sub

Private Sub cboCharity_Change()
    modTCPPv2.gCharityFilter = NzCombo(cboCharity.value, "(All)")
    RefreshDashboard
End Sub

Private Sub cmdIncomeDues_Click()
    OpenEntry "Income", "Dues"
End Sub

Private Sub cmdIncomeDonation_Click()
    OpenEntry "Income", "Donation"
End Sub

Private Sub cmdIncomeEvent_Click()
    OpenEntry "Income", "Event"
End Sub

Private Sub cmdExpense_Click()
    OpenEntry "Expense", "Operations"
End Sub

Private Sub cmdReimbursement_Click()
    OpenEntry "Reimbursement", "Operations"
End Sub

Private Sub cmdAttachFixReceipt_Click()
    frmReceiptInfo.InitForMonth modTCPPv2.gMonthKey
    frmReceiptInfo.Show vbModal
    RefreshDashboard
End Sub

Private Sub cmdReconcileMonth_Click()
    frmReconcile.InitForMonth modTCPPv2.gMonthKey
    frmReconcile.Show vbModal
    RefreshDashboard
End Sub

Private Sub cmdCloseMonth_Click()
    frmCloseMonth.InitForMonth modTCPPv2.gMonthKey
    frmCloseMonth.Show vbModal
    RefreshDashboard
End Sub

Private Sub cmdMonthlyBoardPacket_Click()
    On Error GoTo EH
    modTCPPv2.GenerateMonthlyPacket modTCPPv2.gMonthKey
    lblStatusLastAction.caption = "Last action: Board packet generated (" & modTCPPv2.gMonthKey & ")"
    RefreshDashboard
    Exit Sub
EH:
    modTCPPv2.HandleError "frmDashboard.BoardPacket", Err, modTCPPv2.gMonthKey
End Sub

Private Sub cmdBudget_Click()
    frmBudget.InitForMonth modTCPPv2.gMonthKey
    frmBudget.Show vbModal
    RefreshDashboard
End Sub

Private Sub cmdManageLists_Click()
    frmManageLists.Show vbModal
    LoadEventList
    LoadCharityList
    RefreshDashboard
End Sub

Private Sub cmdMembers_Click()
    frmMembers.Show vbModal
    RefreshDashboard
End Sub

Private Sub cmdMinutes_Click()
    frmMinutesHub.Show vbModal
End Sub

Private Sub cmdAgenda_Click()
    frmAgenda.Show vbModal
End Sub

Private Sub cmdImports_Click()
    frmImports.Show vbModal
End Sub

Private Sub cmdSelfTest_Click()
    modTCPPv2.RunSelfTest
End Sub

Private Sub cmdFixSelected_Click()
    If lstExceptions.ListIndex < 0 Then Exit Sub

    Dim issue As String
    issue = CStr(lstExceptions.List(lstExceptions.ListIndex, 0))

    frmFixIssues.Init issue, modTCPPv2.gMonthKey
    frmFixIssues.Show vbModal
    RefreshDashboard
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'----------------------
' Internal helpers
'----------------------

Private Sub OpenEntry(ByVal txnType As String, ByVal txnDetail As String)
    frmEntry.Init txnType, txnDetail, modTCPPv2.gMonthKey
    frmEntry.Show vbModal
    RefreshDashboard
End Sub

Private Sub RefreshDashboard()
    Dim monthKey As String: monthKey = NzCombo(cboMonth.value, Format(Date, "yyyy-mm"))
    modTCPPv2.gMonthKey = monthKey
    modTCPPv2.gEventFilter = NzCombo(cboEvent.value, "(All)")
    modTCPPv2.gCharityFilter = NzCombo(cboCharity.value, "(All)")

    Dim unc As Long, mr As Long, mAmt As Double
    Dim totalIncome As Double, totalExpense As Double, netChange As Double
    Dim charityRaised As Double, charityPaid As Double, charityHeld As Double
    Dim budgetVarMonth As Double, budgetVarYTD As Double
    modTCPPv2.GetExceptionCounts monthKey, unc, mr, mAmt
    modTCPPv2.GetDashboardMetrics monthKey, modTCPPv2.gEventFilter, modTCPPv2.gCharityFilter, _
        totalIncome, totalExpense, netChange, mr, mAmt, unc, charityRaised, charityPaid, charityHeld, _
        budgetVarMonth, budgetVarYTD

    Dim reconOk As Boolean: reconOk = modTCPPv2.IsReconOk(monthKey)
    Dim closed As Boolean: closed = modTCPPv2.IsMonthClosed(monthKey)

    lblStatusRecon.caption = "Reconciled: " & IIf(reconOk, "YES", "NO")
    lblStatusClosed.caption = "Closed: " & IIf(closed, "YES", "NO")
    lblStatusUncategorized.caption = "Uncategorized: " & CStr(unc)
    lblStatusMissingReceipts.caption = "Missing receipts: " & CStr(mr) & " ($" & Format(mAmt, "0.00") & ")"
    lblStatusCharityHeld.caption = "Charity held (YTD): $" & Format(charityHeld, "0.00")
    lblStatusBudgetVarYTD.caption = "Budget variance (YTD $): $" & Format(budgetVarYTD, "0.00")
    lblStatusTotalIncome.caption = "Total Income (Net): $" & Format(totalIncome, "0.00")
    lblStatusTotalExpense.caption = "Total Expenses (Net): $" & Format(totalExpense, "0.00")
    lblStatusNetChange.caption = "Net Change: $" & Format(netChange, "0.00")
    lblStatusCharityRaised.caption = "Charity Raised (Month): $" & Format(charityRaised, "0.00")
    lblStatusCharityPaid.caption = "Charity Paid (Month): $" & Format(charityPaid, "0.00")
    lblStatusBudgetVarMonth.caption = "Budget variance (Month $): $" & Format(budgetVarMonth, "0.00")

    BuildExceptionsList monthKey, unc, mr, mAmt, reconOk, closed

    ApplyStatusColors reconOk, closed, unc, mr
End Sub

Private Sub ApplyStatusColors(ByVal reconOk As Boolean, ByVal closed As Boolean, ByVal unc As Long, ByVal mr As Long)
    ApplyStatusColor lblStatusRecon, reconOk
    ApplyStatusColor lblStatusClosed, closed
    ApplyStatusColor lblStatusUncategorized, (unc = 0)
    ApplyStatusColor lblStatusMissingReceipts, (mr = 0)
End Sub

Private Sub ApplyStatusColor(ByVal lbl As MSForms.label, ByVal ok As Boolean)
    If ok Then
        lbl.BackColor = RGB(198, 239, 206)
        lbl.ForeColor = RGB(0, 97, 0)
    Else
        lbl.BackColor = RGB(255, 199, 206)
        lbl.ForeColor = RGB(156, 0, 6)
    End If
End Sub

Private Sub BuildExceptionsList(ByVal monthKey As String, ByVal unc As Long, ByVal mr As Long, ByVal mAmt As Double, ByVal reconOk As Boolean, ByVal closed As Boolean)
    Dim charityImbalance As Boolean, budgetOverrun As Boolean
    Dim msg As String
    msg = modTCPPv2.GateCheckMonth(monthKey, unc, mr, mAmt, reconOk, charityImbalance, budgetOverrun)

    lstExceptions.Clear
    If unc > 0 Then AddExceptionRow "Uncategorized", unc
    If mr > 0 Then AddExceptionRow "Missing Receipt", mr
    If Not reconOk Then AddExceptionRow "Not Reconciled", 1
    If Not closed Then AddExceptionRow "Not Closed", 1
    If charityImbalance Then AddExceptionRow "Charity imbalance", 1
    If budgetOverrun Then AddExceptionRow "Budget overrun", 1
End Sub

Private Sub AddExceptionRow(ByVal label As String, ByVal count As Long)
    lstExceptions.AddItem label
    lstExceptions.List(lstExceptions.ListCount - 1, 1) = CStr(count)
End Sub

Private Sub LoadMonthList()
    cboMonth.Clear
    Dim i As Long
    Dim d As Date
    d = DateSerial(Year(Date), Month(Date), 1)
    For i = -12 To 12
        cboMonth.AddItem Format(DateAdd("m", i, d), "yyyy-mm")
    Next i
End Sub

Private Sub LoadEventList()
    cboEvent.Clear
    cboEvent.AddItem "(All)"
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblEventsList")
    If Not lo.DataBodyRange Is Nothing Then
        Dim c As Range
        For Each c In lo.ListColumns(1).DataBodyRange.Cells
            cboEvent.AddItem CStr(c.value)
        Next c
    End If
End Sub

Private Sub LoadCharityList()
    cboCharity.Clear
    cboCharity.AddItem "(All)"
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblCharities")
    If Not lo.DataBodyRange Is Nothing Then
        Dim c As Range
        For Each c In lo.ListColumns(1).DataBodyRange.Cells
            cboCharity.AddItem CStr(c.value)
        Next c
    End If
End Sub

Private Function NzCombo(ByVal v As Variant, ByVal fallback As String) As String
    If IsNull(v) Or Len(Trim$(CStr(v))) = 0 Then
        NzCombo = fallback
    Else
        NzCombo = CStr(v)
    End If
End Function


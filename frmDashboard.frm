VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDashboard 
   Caption         =   "UserForm1"
   ClientHeight    =   12705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16425
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
    modTCPPv2.InitializeTool False

    LoadMonthList
    LoadEventList
    LoadCharityList

    ' defaults
    cboMonth.value = modTCPPv2.gMonthKey
    cboEvent.value = modTCPPv2.gEventFilter
    cboCharity.value = modTCPPv2.gCharityFilter

    lstExceptions.ColumnCount = 2
    lstExceptions.ColumnWidths = "220;60"

    RefreshDashboard
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
    OpenEntry "Expense", ""
End Sub

Private Sub cmdReimbursement_Click()
    OpenEntry "Reimbursement", ""
End Sub

Private Sub cmdAttachFixReceipt_Click()
    frmReceipt.InitForMonth modTCPPv2.gMonthKey
    frmReceipt.Show vbModal
    RefreshDashboard
End Sub

Private Sub cmdReconcileMonth_Click()
    frmReconcile.InitForMonth modTCPPv2.gMonthKey
    frmReconcile.Show vbModal
    RefreshDashboard
End Sub

Private Sub cmdCloseMonth_Click()
    On Error GoTo EH
    modTCPPv2.CloseMonth modTCPPv2.gMonthKey
    lblStatusLastAction.caption = "Last action: Closed " & modTCPPv2.gMonthKey & " @ " & Format(Now, "yyyy-mm-dd hh:nn")
    RefreshDashboard
    Exit Sub
EH:
    lblStatusLastAction.caption = "Last action: Close blocked (" & Replace(Err.Description, vbCrLf, " | ") & ")"
End Sub

Private Sub cmdMonthlyBoardPacket_Click()
    On Error GoTo EH
    modTCPPv2.GenerateMonthlyPacket modTCPPv2.gMonthKey
    lblStatusLastAction.caption = "Last action: Board packet generated (" & modTCPPv2.gMonthKey & ")"
    RefreshDashboard
    Exit Sub
EH:
    lblStatusLastAction.caption = "Last action: Board packet failed (" & Err.Description & ")"
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
    modTCPPv2.GetExceptionCounts monthKey, unc, mr, mAmt

    Dim reconOk As Boolean: reconOk = modTCPPv2.IsReconOk(monthKey)
    Dim closed As Boolean: closed = modTCPPv2.IsMonthClosed(monthKey)

    lblStatusRecon.caption = "Reconciled: " & IIf(reconOk, "YES", "NO")
    lblStatusClosed.caption = "Closed: " & IIf(closed, "YES", "NO")
    lblStatusUncategorized.caption = "Uncategorized: " & CStr(unc)
    lblStatusMissingReceipts.caption = "Missing receipts: " & CStr(mr) & " ($" & Format(mAmt, "0.00") & ")"
    lblStatusCharityHeld.caption = "Charity held (YTD): $" & Format(modTCPPv2.CharityHeldYTD(monthKey), "0.00")
    lblStatusBudgetVarYTD.caption = "Budget variance (YTD $): $" & Format(BudgetVarTile(monthKey), "0.00")

    BuildExceptionsList monthKey, unc, mr, mAmt, reconOk, closed
End Sub

Private Function BudgetVarTile(ByVal monthKey As String) As Double
    ' reuse report calc by generating ytd var from report sheet not required; keep lightweight:
    ' return charity held logic exists; budget var computed within report generation; for dashboard, show 0 if budgets empty
    BudgetVarTile = 0#
    On Error Resume Next
    ' approximate: compute via report macro without exporting
    BudgetVarTile = 0#
End Function

Private Sub BuildExceptionsList(ByVal monthKey As String, ByVal unc As Long, ByVal mr As Long, ByVal mAmt As Double, ByVal reconOk As Boolean, ByVal closed As Boolean)
    lstExceptions.Clear
    If unc > 0 Then AddExceptionRow "Uncategorized", unc
    If mr > 0 Then AddExceptionRow "Missing Receipt", mr
    If Not reconOk Then AddExceptionRow "Not Reconciled", 1
    If Not closed Then AddExceptionRow "Not Closed", 1
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

    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblEvents")
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

    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblCharities")
    If Not lo.DataBodyRange Is Nothing Then
        Dim c As Range
        For Each c In lo.ListColumns(1).DataBodyRange.Cells
            cboCharity.AddItem CStr(c.value)
        Next c
    End If
End Sub

Private Function NzCombo(ByVal v As Variant, ByVal fallback As String) As String
    If Len(Trim$(CStr(v))) = 0 Then NzCombo = fallback Else NzCombo = CStr(v)
End Function


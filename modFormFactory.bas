Attribute VB_Name = "modFormFactory"
Option Explicit

'========================
' Form Factory (rebuilds missing controls when .frx files are unavailable)
' Requires: Trust access to the VBA project object model
'========================

Public Sub RebuildAllForms()
    On Error GoTo EH
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    BuildFrmDashboard vbProj
    BuildFrmEntry vbProj
    BuildFrmReceiptInfo vbProj
    BuildFrmCloseMonth vbProj
    BuildFrmMembers vbProj
    BuildFrmMinutesHub vbProj
    BuildFrmAgenda vbProj
    BuildFrmAttendance vbProj
    BuildFrmAgendaLines vbProj
    BuildFrmActionItems vbProj
    BuildFrmImports vbProj
    BuildFrmManageLists vbProj
    BuildFrmReceipt vbProj
    BuildFrmReconcile vbProj
    BuildFrmBudget vbProj
    BuildFrmFixIssues vbProj

    MsgBox "Form rebuild complete. Save the workbook to generate .frx files.", vbInformation, "Form Factory"
    Exit Sub
EH:
    MsgBox "Form rebuild failed. Ensure 'Trust access to the VBA project object model' is enabled." & vbCrLf & Err.Description, vbExclamation, "Form Factory"
End Sub

Private Sub BuildFrmDashboard(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmDashboard")
    ClearControls frm
    AddCombo frm, "cboMonth", 12, 12, 120, 20
    AddCombo frm, "cboEvent", 140, 12, 140, 20
    AddCombo frm, "cboCharity", 290, 12, 140, 20
    AddListBox frm, "lstExceptions", 12, 40, 260, 120

    AddButton frm, "cmdRefreshDashboard", "Refresh", 440, 12, 80, 22
    AddButton frm, "cmdIncomeDues", "Income Dues", 12, 170, 90, 22
    AddButton frm, "cmdIncomeDonation", "Income Donation", 108, 170, 110, 22
    AddButton frm, "cmdIncomeEvent", "Income Event", 224, 170, 90, 22
    AddButton frm, "cmdExpense", "Expense", 320, 170, 80, 22
    AddButton frm, "cmdReimbursement", "Reimb", 406, 170, 80, 22
    AddButton frm, "cmdAttachFixReceipt", "Receipts", 12, 198, 80, 22
    AddButton frm, "cmdReconcileMonth", "Reconcile", 98, 198, 80, 22
    AddButton frm, "cmdCloseMonth", "Close Month", 184, 198, 90, 22
    AddButton frm, "cmdMonthlyBoardPacket", "Board Packet", 280, 198, 90, 22
    AddButton frm, "cmdBudget", "Budget", 376, 198, 70, 22
    AddButton frm, "cmdManageLists", "Lists", 452, 198, 60, 22

    AddButton frm, "cmdMembers", "Members", 12, 226, 80, 22
    AddButton frm, "cmdMinutes", "Minutes", 98, 226, 80, 22
    AddButton frm, "cmdAgenda", "Agenda", 184, 226, 80, 22
    AddButton frm, "cmdImports", "Imports", 270, 226, 80, 22
    AddButton frm, "cmdSelfTest", "SelfTest", 356, 226, 70, 22
    AddButton frm, "cmdFixSelected", "Fix Selected", 432, 226, 80, 22
    AddButton frm, "cmdExit", "Exit", 518, 226, 60, 22

    AddLabel frm, "lblStatusRecon", "Reconciled:", 280, 40, 160, 18
    AddLabel frm, "lblStatusClosed", "Closed:", 280, 60, 160, 18
    AddLabel frm, "lblStatusUncategorized", "Uncategorized:", 280, 80, 160, 18
    AddLabel frm, "lblStatusMissingReceipts", "Missing receipts:", 280, 100, 200, 18
    AddLabel frm, "lblStatusCharityHeld", "Charity held:", 280, 120, 200, 18
    AddLabel frm, "lblStatusBudgetVarYTD", "Budget YTD:", 280, 140, 200, 18
    AddLabel frm, "lblStatusTotalIncome", "Total Income:", 12, 260, 200, 18
    AddLabel frm, "lblStatusTotalExpense", "Total Expense:", 12, 280, 200, 18
    AddLabel frm, "lblStatusNetChange", "Net Change:", 12, 300, 200, 18
    AddLabel frm, "lblStatusCharityRaised", "Charity Raised:", 220, 260, 220, 18
    AddLabel frm, "lblStatusCharityPaid", "Charity Paid:", 220, 280, 220, 18
    AddLabel frm, "lblStatusBudgetVarMonth", "Budget Month:", 220, 300, 220, 18
    AddLabel frm, "lblStatusLastAction", "Last action:", 12, 330, 400, 18
End Sub

Private Sub BuildFrmEntry(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmEntry")
    ClearControls frm
    AddLabel frm, "lblTxnID", "", 12, 12, 200, 18
    AddLabel frm, "lblNet", "Net:", 12, 36, 200, 18
    AddText frm, "txtDate", 12, 60, 100, 20
    AddCombo frm, "cboTxnType", 120, 60, 120, 20
    AddCombo frm, "cboCategory", 250, 60, 120, 20
    AddCombo frm, "cboEvent", 380, 60, 120, 20
    AddCombo frm, "cboCharity", 510, 60, 120, 20
    AddText frm, "txtGross", 12, 90, 80, 20
    AddText frm, "txtFees", 100, 90, 80, 20
    AddCombo frm, "cboPaymentMethod", 190, 90, 120, 20
    AddCombo frm, "cboSourceType", 320, 90, 120, 20
    AddText frm, "txtPayeeSource", 450, 90, 180, 20
    AddText frm, "txtMemberName", 12, 120, 150, 20
    AddText frm, "txtMemberEmail", 170, 120, 200, 20
    AddText frm, "txtMemo", 12, 150, 300, 40
    AddCheckbox frm, "chkReceiptRequired", "Receipt Required", 320, 150, 120, 18
    AddButton frm, "cmdSave", "Save", 12, 200, 60, 22
    AddButton frm, "cmdSaveAndAttach", "Save+Receipt", 80, 200, 100, 22
    AddButton frm, "cmdCancel", "Cancel", 190, 200, 60, 22
End Sub

Private Sub BuildFrmReceiptInfo(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmReceiptInfo")
    ClearControls frm
    AddCombo frm, "cboMonth", 12, 12, 100, 20
    AddText frm, "txtTxnId", 120, 12, 120, 20
    AddText frm, "txtVendor", 250, 12, 150, 20
    AddButton frm, "cmdSearch", "Search", 410, 12, 60, 20
    AddListBox frm, "lstTxns", 12, 40, 520, 140
    AddButton frm, "cmdRecord", "Record", 12, 190, 60, 22
    AddButton frm, "cmdWaive", "Waive", 80, 190, 60, 22
    AddText frm, "txtWaiveReason", 150, 190, 220, 20
    AddButton frm, "cmdClose", "Close", 470, 190, 60, 22
End Sub

Private Sub BuildFrmCloseMonth(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmCloseMonth")
    ClearControls frm
    AddCombo frm, "cboMonth", 12, 12, 100, 20
    AddButton frm, "cmdCheck", "Check Gates", 120, 12, 90, 20
    AddButton frm, "cmdCloseMonth", "Close Month", 220, 12, 90, 20
    AddButton frm, "cmdClose", "Close", 320, 12, 60, 20
    AddText frm, "txtGateSummary", 12, 40, 520, 120
    frm.Controls("txtGateSummary").MultiLine = True
End Sub

Private Sub BuildFrmMembers(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmMembers")
    ClearControls frm
    AddText frm, "txtName", 12, 12, 140, 20
    AddText frm, "txtEmail", 160, 12, 180, 20
    AddCombo frm, "cboMembershipType", 350, 12, 120, 20
    AddCheckbox frm, "chkDuesPaid", "Dues Paid", 480, 12, 80, 18
    AddText frm, "txtPaidDate", 12, 40, 120, 20
    AddText frm, "txtDuesAmount", 140, 40, 80, 20
    AddText frm, "txtJoinedDate", 230, 40, 100, 20
    AddText frm, "txtNotes", 340, 40, 180, 20
    AddText frm, "txtRenewalDate", 12, 68, 120, 20
    AddButton frm, "cmdSave", "Save", 540, 12, 60, 20
    AddButton frm, "cmdSearch", "Search", 540, 40, 60, 20
    AddButton frm, "cmdExportReport", "Export", 610, 40, 60, 20
    AddButton frm, "cmdClose", "Close", 610, 12, 60, 20
    AddListBox frm, "lstMembers", 12, 100, 660, 140
End Sub

Private Sub BuildFrmMinutesHub(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmMinutesHub")
    ClearControls frm
    AddButton frm, "cmdNewMeeting", "New", 12, 12, 50, 20
    AddButton frm, "cmdExportPdf", "Export", 70, 12, 60, 20
    AddButton frm, "cmdOpenDoc", "Open DOC", 140, 12, 70, 20
    AddButton frm, "cmdOpenPdf", "Open PDF", 220, 12, 70, 20
    AddButton frm, "cmdAttendance", "Attendance", 300, 12, 80, 20
    AddButton frm, "cmdAgendaLines", "Agenda Lines", 390, 12, 90, 20
    AddButton frm, "cmdActionItems", "Action Items", 490, 12, 90, 20
    AddText frm, "txtSearch", 590, 12, 100, 20
    AddButton frm, "cmdSearch", "Search", 700, 12, 60, 20
    AddButton frm, "cmdClose", "Close", 770, 12, 60, 20
    AddListBox frm, "lstMeetings", 12, 40, 820, 160
End Sub

Private Sub BuildFrmAgenda(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmAgenda")
    ClearControls frm
    AddButton frm, "cmdNewAgenda", "New", 12, 12, 50, 20
    AddButton frm, "cmdExportPdf", "Export", 70, 12, 60, 20
    AddButton frm, "cmdOpenDoc", "Open DOC", 140, 12, 70, 20
    AddButton frm, "cmdOpenPdf", "Open PDF", 220, 12, 70, 20
    AddText frm, "txtSearch", 300, 12, 120, 20
    AddButton frm, "cmdSearch", "Search", 430, 12, 60, 20
    AddButton frm, "cmdClose", "Close", 500, 12, 60, 20
    AddListBox frm, "lstAgendas", 12, 40, 560, 160
End Sub

Private Sub BuildFrmAttendance(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmAttendance")
    ClearControls frm
    AddText frm, "txtPersonName", 12, 12, 140, 20
    AddText frm, "txtRole", 160, 12, 100, 20
    AddCheckbox frm, "chkPresent", "Present", 270, 12, 70, 18
    AddButton frm, "cmdAdd", "Add", 350, 12, 50, 20
    AddButton frm, "cmdDelete", "Delete", 410, 12, 50, 20
    AddListBox frm, "lstAttendance", 12, 40, 460, 140
End Sub

Private Sub BuildFrmAgendaLines(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmAgendaLines")
    ClearControls frm
    AddText frm, "txtLineTime", 12, 12, 80, 20
    AddText frm, "txtTopic", 100, 12, 140, 20
    AddText frm, "txtOwner", 250, 12, 100, 20
    AddText frm, "txtActionItem", 360, 12, 140, 20
    AddButton frm, "cmdAdd", "Add", 510, 12, 50, 20
    AddButton frm, "cmdDelete", "Delete", 570, 12, 50, 20
    AddListBox frm, "lstLines", 12, 40, 620, 140
End Sub

Private Sub BuildFrmActionItems(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmActionItems")
    ClearControls frm
    AddText frm, "txtActionItem", 12, 12, 160, 20
    AddText frm, "txtOwner", 180, 12, 100, 20
    AddText frm, "txtDueDate", 290, 12, 100, 20
    AddCombo frm, "cboStatus", 400, 12, 80, 20
    AddButton frm, "cmdAdd", "Add", 490, 12, 50, 20
    AddButton frm, "cmdDelete", "Delete", 550, 12, 50, 20
    AddListBox frm, "lstActions", 12, 40, 600, 140
End Sub

Private Sub BuildFrmImports(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmImports")
    ClearControls frm
    AddCombo frm, "cboSource", 12, 12, 80, 20
    AddText frm, "txtFilePath", 100, 12, 220, 20
    AddButton frm, "cmdBrowse", "Browse", 330, 12, 60, 20
    AddButton frm, "cmdImport", "Import", 12, 40, 60, 20
    AddButton frm, "cmdClose", "Close", 80, 40, 60, 20
End Sub

Private Sub BuildFrmManageLists(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmManageLists")
    ClearControls frm
    AddListBox frm, "lstEvents", 12, 40, 140, 80
    AddText frm, "txtNewEvent", 12, 12, 100, 20
    AddButton frm, "cmdAddEvent", "Add", 120, 12, 40, 20
    AddButton frm, "cmdRemoveEvent", "Remove", 120, 40, 50, 20

    AddListBox frm, "lstCharities", 180, 40, 140, 80
    AddText frm, "txtNewCharity", 180, 12, 100, 20
    AddButton frm, "cmdAddCharity", "Add", 288, 12, 40, 20
    AddButton frm, "cmdRemoveCharity", "Remove", 288, 40, 50, 20

    AddListBox frm, "lstCOA", 350, 40, 140, 80
    AddText frm, "txtNewCOA", 350, 12, 100, 20
    AddButton frm, "cmdAddCOA", "Add", 458, 12, 40, 20
    AddButton frm, "cmdRemoveCOA", "Remove", 458, 40, 50, 20

    AddListBox frm, "lstPaymentMethods", 12, 140, 140, 80
    AddText frm, "txtNewPaymentMethod", 12, 112, 100, 20
    AddButton frm, "cmdAddPaymentMethod", "Add", 120, 112, 40, 20
    AddButton frm, "cmdRemovePaymentMethod", "Remove", 120, 140, 50, 20

    AddListBox frm, "lstTxnTypes", 180, 140, 140, 80
    AddText frm, "txtNewTxnType", 180, 112, 100, 20
    AddButton frm, "cmdAddTxnType", "Add", 288, 112, 40, 20
    AddButton frm, "cmdRemoveTxnType", "Remove", 288, 140, 50, 20

    AddListBox frm, "lstTxnSubtypes", 350, 140, 140, 80
    AddText frm, "txtNewTxnSubtype", 350, 112, 100, 20
    AddButton frm, "cmdAddTxnSubtype", "Add", 458, 112, 40, 20
    AddButton frm, "cmdRemoveTxnSubtype", "Remove", 458, 140, 50, 20

    AddListBox frm, "lstMembershipTypes", 12, 240, 140, 80
    AddText frm, "txtNewMembershipType", 12, 212, 100, 20
    AddButton frm, "cmdAddMembershipType", "Add", 120, 212, 40, 20
    AddButton frm, "cmdRemoveMembershipType", "Remove", 120, 240, 50, 20

    AddListBox frm, "lstBoardRoster", 180, 240, 140, 80
    AddText frm, "txtBoardName", 180, 212, 80, 20
    AddText frm, "txtBoardRole", 266, 212, 80, 20
    AddButton frm, "cmdAddBoard", "Add", 352, 212, 40, 20
    AddButton frm, "cmdRemoveBoard", "Remove", 352, 240, 50, 20

    AddButton frm, "cmdClose", "Close", 450, 240, 50, 20
End Sub

Private Sub BuildFrmReceipt(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmReceipt")
    ClearControls frm
    AddCombo frm, "cboMonth", 12, 12, 100, 20
    AddListBox frm, "lstTxns", 12, 40, 400, 120
    AddButton frm, "cmdAttach", "Record", 12, 170, 60, 20
    AddButton frm, "cmdWaive", "Waive", 80, 170, 60, 20
    AddText frm, "txtWaiveReason", 150, 170, 180, 20
    AddButton frm, "cmdOpenFile", "Open", 340, 170, 60, 20
    AddButton frm, "cmdClose", "Close", 410, 170, 60, 20
End Sub

Private Sub BuildFrmReconcile(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmReconcile")
    ClearControls frm
    AddLabel frm, "lblMonth", "Month:", 12, 12, 100, 18
    AddText frm, "txtBeginningBalance", 12, 40, 100, 20
    AddText frm, "txtEndingBalance", 120, 40, 100, 20
    AddLabel frm, "lblLedgerDeposits", "Deposits:", 12, 70, 200, 18
    AddLabel frm, "lblLedgerWithdrawals", "Withdrawals:", 12, 90, 200, 18
    AddLabel frm, "lblExpectedEnding", "Expected:", 12, 110, 200, 18
    AddLabel frm, "lblDifference", "Difference:", 12, 130, 200, 18
    AddButton frm, "cmdCompute", "Compute", 12, 160, 60, 20
    AddButton frm, "cmdSave", "Save", 80, 160, 60, 20
    AddButton frm, "cmdClose", "Close", 150, 160, 60, 20
End Sub

Private Sub BuildFrmBudget(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmBudget")
    ClearControls frm
    AddCombo frm, "cboMonth", 12, 12, 100, 20
    AddText frm, "txtBudgetAdministrative", 12, 40, 100, 20
    AddText frm, "txtBudgetPrograms", 12, 65, 100, 20
    AddText frm, "txtBudgetFundraising", 12, 90, 100, 20
    AddText frm, "txtBudgetMarketing", 12, 115, 100, 20
    AddText frm, "txtBudgetTravel", 12, 140, 100, 20
    AddText frm, "txtBudgetServices", 12, 165, 100, 20
    AddText frm, "txtBudgetMisc", 12, 190, 100, 20
    AddButton frm, "cmdLoad", "Load", 120, 12, 60, 20
    AddButton frm, "cmdSave", "Save", 120, 40, 60, 20
    AddButton frm, "cmdClose", "Close", 120, 70, 60, 20
End Sub

Private Sub BuildFrmFixIssues(ByVal vbProj As Object)
    Dim frm As Object: Set frm = GetOrCreateForm(vbProj, "frmFixIssues")
    ClearControls frm
    AddCombo frm, "cboIssueType", 12, 12, 120, 20
    AddCombo frm, "cboMonth", 140, 12, 100, 20
    AddCombo frm, "cboCategory", 250, 12, 120, 20
    AddCombo frm, "cboEvent", 380, 12, 120, 20
    AddCombo frm, "cboCharity", 510, 12, 120, 20
    AddListBox frm, "lstIssues", 12, 40, 400, 120
    AddCheckbox frm, "chkReceiptRequired", "Receipt Required", 420, 40, 120, 18
    AddText frm, "txtWaiveReason", 420, 60, 200, 20
    AddButton frm, "cmdSaveChanges", "Save", 420, 90, 60, 20
    AddButton frm, "cmdAttachReceipt", "Record", 490, 90, 60, 20
    AddButton frm, "cmdWaiveReceipt", "Waive", 560, 90, 60, 20
    AddButton frm, "cmdClose", "Close", 630, 90, 60, 20
End Sub

Private Function GetOrCreateForm(ByVal vbProj As Object, ByVal formName As String) As Object
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = vbProj.VBComponents.Item(formName)
    On Error GoTo 0

    If vbComp Is Nothing Then
        Set vbComp = vbProj.VBComponents.Add(3)
        vbComp.Name = formName
    End If
    Set GetOrCreateForm = vbComp.Designer
End Function

Private Sub ClearControls(ByVal frm As Object)
    Dim ctl As Object
    For Each ctl In frm.Controls
        frm.Controls.Remove ctl.Name
    Next ctl
End Sub

Private Sub AddButton(ByVal frm As Object, ByVal name As String, ByVal caption As String, _
                      ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single)
    Dim ctl As Object
    Set ctl = frm.Controls.Add("Forms.CommandButton.1", name, True)
    ctl.Caption = caption
    ctl.Left = left
    ctl.Top = top
    ctl.Width = width
    ctl.Height = height
End Sub

Private Sub AddLabel(ByVal frm As Object, ByVal name As String, ByVal caption As String, _
                     ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single)
    Dim ctl As Object
    Set ctl = frm.Controls.Add("Forms.Label.1", name, True)
    ctl.Caption = caption
    ctl.Left = left
    ctl.Top = top
    ctl.Width = width
    ctl.Height = height
End Sub

Private Sub AddText(ByVal frm As Object, ByVal name As String, _
                    ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single)
    Dim ctl As Object
    Set ctl = frm.Controls.Add("Forms.TextBox.1", name, True)
    ctl.Left = left
    ctl.Top = top
    ctl.Width = width
    ctl.Height = height
End Sub

Private Sub AddCombo(ByVal frm As Object, ByVal name As String, _
                     ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single)
    Dim ctl As Object
    Set ctl = frm.Controls.Add("Forms.ComboBox.1", name, True)
    ctl.Left = left
    ctl.Top = top
    ctl.Width = width
    ctl.Height = height
End Sub

Private Sub AddListBox(ByVal frm As Object, ByVal name As String, _
                       ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single)
    Dim ctl As Object
    Set ctl = frm.Controls.Add("Forms.ListBox.1", name, True)
    ctl.Left = left
    ctl.Top = top
    ctl.Width = width
    ctl.Height = height
End Sub

Private Sub AddCheckbox(ByVal frm As Object, ByVal name As String, ByVal caption As String, _
                        ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single)
    Dim ctl As Object
    Set ctl = frm.Controls.Add("Forms.CheckBox.1", name, True)
    ctl.Caption = caption
    ctl.Left = left
    ctl.Top = top
    ctl.Width = width
    ctl.Height = height
End Sub

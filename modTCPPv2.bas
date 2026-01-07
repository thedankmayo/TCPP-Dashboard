Attribute VB_Name = "modTCPPv2"
Option Explicit

'========================
' Acceptance Tests
'========================
' 1) Create Expense -> record receipt metadata -> reconcile totals -> close month -> generate board packet PDF + archive snapshot.
' 2) Attempt to close month with missing receipt -> blocked.
' 3) Attempt to edit closed-month transaction -> blocked.
' 4) Create meeting -> Word doc created from template and opened -> export PDF -> paths saved -> meeting appears in archive list.
' 5) Add members -> mark dues paid -> renewal date auto-calculated -> search works.
' 6) Force an error (e.g., missing folder) -> error logged in tblErrorLog with procedure and context.

'========================
' v2: Treasurer + Secretary unified tool
'========================

'--- Sheet names
Private Const SH_HOME As String = "HOME"
Private Const SH_LEDGER As String = "DATA_Ledger"
Private Const SH_RECEIPTS As String = "DATA_Receipts"
Private Const SH_LOOKUPS As String = "DATA_Lookups"
Private Const SH_BUDGET As String = "DATA_Budget"
Private Const SH_MONTHSTATUS As String = "DATA_MonthStatus"
Private Const SH_MEMBERS As String = "DATA_Members"
Private Const SH_MEETINGS As String = "DATA_Meetings"
Private Const SH_ATTENDANCE As String = "DATA_Attendance"
Private Const SH_MINUTES_LINES As String = "DATA_MinutesAgenda"
Private Const SH_ACTION_ITEMS As String = "DATA_ActionItems"
Private Const SH_AGENDA As String = "DATA_Agenda"
Private Const SH_IMPORTS As String = "DATA_Imports"
Private Const SH_AUDIT As String = "DATA_Audit"
Private Const SH_ERRORLOG As String = "DATA_ErrorLog"
Private Const SH_COMPLIANCE As String = "DATA_Compliance"
Private Const SH_REPORT As String = "RPT_Monthly"
Private Const SH_COMPLIANCE_REPORT As String = "RPT_Compliance"
Private Const SH_ARCHIVE As String = "ARCH_BoardPackets"

'--- Table names
Private Const T_LEDGER As String = "tblLedger"
Private Const T_RECEIPTS As String = "tblReceipts"
Private Const T_BUDGET As String = "tblBudget"
Private Const T_MONTHSTATUS As String = "tblMonthStatus"
Private Const T_MEMBERS As String = "tblMembers"
Private Const T_MEETINGS As String = "tblMeetings"
Private Const T_ATTENDANCE As String = "tblMeetingAttendance"
Private Const T_MINUTES_LINES As String = "tblMinutesAgendaLines"
Private Const T_ACTION_ITEMS As String = "tblActionItems"
Private Const T_AGENDA As String = "tblAgenda"
Private Const T_EVENT_ATTENDANCE As String = "tblEventAttendance"
Private Const T_IMPORTLOG As String = "tblImportLog"
Private Const T_ZEFFY_RAW As String = "tblZeffyRaw"
Private Const T_BLAZE_RAW As String = "tblBlazeRaw"
Private Const T_IMPORTMAP_ZEFFY As String = "tblImportMap_Zeffy"
Private Const T_IMPORTMAP_BLAZE As String = "tblImportMap_Blaze"
Private Const T_ERRORLOG As String = "tblErrorLog"
Private Const T_AUDIT As String = "tblAuditLog"
Private Const T_COMPLIANCE_EVENTS As String = "tblEvents"
Private Const T_COMPLIANCE_VENDORS As String = "tblVendorsPartners"
Private Const T_COMPLIANCE_VOLUNTEERS As String = "tblVolunteers"
Private Const T_COMPLIANCE_GIVEAWAYS As String = "tblGiveaways"
Private Const T_COMPLIANCE_REFUNDS As String = "tblRefundRequests"
Private Const T_COMPLIANCE_INCIDENTS As String = "tblIncidentsComplaints"
Private Const T_COMPLIANCE_OUTCOMES As String = "tblOutcomes"
Private Const T_COMPLIANCE_APPEALS As String = "tblAppeals"
Private Const T_COMPLIANCE_ACKS As String = "tblPolicyAcks"
Private Const T_COMPLIANCE_PAYMENT_ADMINS As String = "tblPaymentAdmins"

'--- Lookup tables
Private Const T_COA As String = "tblCOA"
Private Const T_TXN_TYPES As String = "tblTxnTypes"
Private Const T_TXN_SUBTYPES As String = "tblTxnSubtypes"
Private Const T_EVENTS_LIST As String = "tblEventsList"
Private Const T_CHARITIES As String = "tblCharities"
Private Const T_PAYMETHOD As String = "tblPaymentMethods"
Private Const T_BOARDROSTER As String = "tblBoardRoster"
Private Const T_MEMBERTYPES As String = "tblMembershipTypes"
Private Const T_MEETING_TYPES As String = "tblMeetingTypes"
Private Const T_CONFIG As String = "tblConfig"

'--- Config keys
Private Const CFG_FISCAL_START_MONTH As String = "FiscalYearStartMonth"
Private Const CFG_APPROVER_NAME As String = "ApproverName"
Private Const CFG_RENEWAL_INTERVAL As String = "RenewalIntervalMonths"
Private Const CFG_STRICT_BUDGET As String = "StrictBudgetGate"
Private Const CFG_RECEIPT_WAIVER_REASON As String = "ReceiptWaiverRequiresReason"
Private Const CFG_PATH_BOARDPACKETS As String = "BoardPacketsFolderRelative"
Private Const CFG_PATH_MINUTES_DOCX As String = "MinutesDocxFolderRelative"
Private Const CFG_PATH_MINUTES_PDF As String = "MinutesPdfFolderRelative"
Private Const CFG_PATH_AGENDA_DOCX As String = "AgendaDocxFolderRelative"
Private Const CFG_PATH_AGENDA_PDF As String = "AgendaPdfFolderRelative"
Private Const CFG_PATH_IMPORTS_ZEFFY As String = "ImportsZeffyFolderRelative"
Private Const CFG_PATH_IMPORTS_BLAZE As String = "ImportsBlazeFolderRelative"
Private Const CFG_PATH_COMPLIANCE_DOCS As String = "ComplianceDocsFolderRelative"
Private Const CFG_PATH_IMPORTS_ARCHIVE As String = "ImportsArchiveFolderRelative"
Private Const CFG_MEETING_NOTICE_REGULAR As String = "RegularMeetingNoticeDays"
Private Const CFG_MEETING_NOTICE_ANNUAL As String = "AnnualMeetingNoticeDays"
Private Const CFG_MEETING_NOTICE_SPECIAL As String = "SpecialMeetingNoticeHours"
Private Const CFG_BOARD_QUORUM As String = "BoardQuorumCount"
Private Const CFG_DUES_INITIAL_DUE_DAYS As String = "DuesInitialDueDaysAfterApproval"
Private Const CFG_DUES_AMOUNT_FULL As String = "DuesAmount_Full"
Private Const CFG_DUES_AMOUNT_ATLARGE As String = "DuesAmount_AtLarge"
Private Const CFG_DUES_SPLIT_FULL As String = "DuesSplitAllowed_Full"
Private Const CFG_DUES_SPLIT_ATLARGE As String = "DuesSplitAllowed_AtLarge"
Private Const CFG_MIN_AGE_FULL As String = "MembershipMinAge_Full"
Private Const CFG_MIN_AGE_ATLARGE As String = "MembershipMinAge_AtLarge"
Private Const CFG_EVENTS_REQ_FULL As String = "FullMemberEventsReqPerCalendarYear"
Private Const CFG_VOLUNTEER_REQ_FULL As String = "FullMemberVolunteerEventsReqPerCalendarYear"
Private Const CFG_RETENTION_MINOR As String = "RetentionYears_MinorIncident"
Private Const CFG_RETENTION_SEVERE As String = "RetentionYears_SevereOrConsent"
Private Const CFG_REFUND_APPROVAL_SELF_OK As String = "RefundAllowSelfApproval"

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

    On Error GoTo EH

    EnsureCoreSheets forceRebuild
    EnsureCoreTables forceRebuild
    SeedLookupsIfEmpty
    EnsureConfigDefaults
    EnsureDefaultFolders
    LockDownWorkbookUI

    If Len(gMonthKey) = 0 Then gMonthKey = Format(Date, "yyyy-mm")
    If Len(gEventFilter) = 0 Then gEventFilter = "(All)"
    If Len(gCharityFilter) = 0 Then gCharityFilter = "(All)"

CleanExit:
    Application.EnableEvents = prevEV
    Application.ScreenUpdating = prevSU
    Exit Sub
EH:
    LogError "InitializeTool", Err, "forceRebuild=" & CStr(forceRebuild)
    Resume CleanExit
End Sub

Public Sub ShowDashboard()
    On Error GoTo EH
    Unload frmDashboard
    frmDashboard.Show vbModal
    Exit Sub
EH:
    HandleError "ShowDashboard", Err, ""
End Sub

Public Sub CleanupOnClose()
    On Error GoTo EH
    Application.DisplayAlerts = True
    Exit Sub
EH:
    HandleError "CleanupOnClose", Err, ""
End Sub

Public Sub RunSelfTest()
    On Error GoTo EH
    Dim msg As String
    msg = SelfTestReport()
    MsgBox msg, vbInformation, "TCPP v2 Self Test"
    Exit Sub
EH:
    HandleError "RunSelfTest", Err, ""
End Sub

'========================
' Core structure
'========================

Private Sub EnsureCoreSheets(ByVal forceRebuild As Boolean)
    EnsureSheet SH_HOME, xlSheetVisible
    EnsureSheet SH_LEDGER, xlSheetVeryHidden
    EnsureSheet SH_RECEIPTS, xlSheetVeryHidden
    EnsureSheet SH_LOOKUPS, xlSheetVeryHidden
    EnsureSheet SH_BUDGET, xlSheetVeryHidden
    EnsureSheet SH_MONTHSTATUS, xlSheetVeryHidden
    EnsureSheet SH_MEMBERS, xlSheetVeryHidden
    EnsureSheet SH_MEETINGS, xlSheetVeryHidden
    EnsureSheet SH_ATTENDANCE, xlSheetVeryHidden
    EnsureSheet SH_MINUTES_LINES, xlSheetVeryHidden
    EnsureSheet SH_ACTION_ITEMS, xlSheetVeryHidden
    EnsureSheet SH_AGENDA, xlSheetVeryHidden
    EnsureSheet SH_IMPORTS, xlSheetVeryHidden
    EnsureSheet SH_AUDIT, xlSheetVeryHidden
    EnsureSheet SH_ERRORLOG, xlSheetVeryHidden
    EnsureSheet SH_COMPLIANCE, xlSheetVeryHidden
    EnsureSheet SH_REPORT, xlSheetVeryHidden
    EnsureSheet SH_COMPLIANCE_REPORT, xlSheetVeryHidden
    EnsureSheet SH_ARCHIVE, xlSheetVeryHidden

    With GetSheet(SH_HOME)
        .Cells.ClearContents
        .Range("A1").value = "TCPP Treasurer + Secretary Dashboard"
        .Range("A2").value = "UserForm hub. Keep this sheet open."
    End With
End Sub

Private Sub EnsureCoreTables(ByVal forceRebuild As Boolean)
    EnsureLookupTables forceRebuild
    EnsureLedgerTable forceRebuild
    EnsureReceiptsTable forceRebuild
    EnsureBudgetTable forceRebuild
    EnsureMonthStatusTable forceRebuild
    EnsureMembersTable forceRebuild
    EnsureMeetingsTables forceRebuild
    EnsureAgendaTable forceRebuild
    EnsureComplianceTables forceRebuild
    EnsureImportsTables forceRebuild
    EnsureAuditTable forceRebuild
    EnsureErrorLogTable forceRebuild
    EnsureReportSheetLayout
    EnsureComplianceReportLayout
    EnsureArchiveTable forceRebuild
End Sub

Private Sub EnsureLookupTables(ByVal forceRebuild As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_LOOKUPS)

    EnsureTable ws, T_COA, Array("Category"), 1, 1, forceRebuild
    EnsureTable ws, T_TXN_TYPES, Array("TxnType"), 1, 3, forceRebuild
    EnsureTable ws, T_TXN_SUBTYPES, Array("TxnSubtype"), 1, 5, forceRebuild
    RenameLookupTable "tblEvents", T_EVENTS_LIST
    EnsureTable ws, T_EVENTS_LIST, Array("Event"), 1, 7, forceRebuild
    EnsureTable ws, T_CHARITIES, Array("Charity"), 1, 9, forceRebuild
    EnsureTable ws, T_PAYMETHOD, Array("PaymentMethod"), 1, 11, forceRebuild
    EnsureTable ws, T_BOARDROSTER, Array("Name", "Role", "ActiveFlag"), 1, 13, forceRebuild
    EnsureTable ws, T_MEMBERTYPES, Array("MembershipType"), 1, 16, forceRebuild
    EnsureTable ws, T_MEETING_TYPES, Array("MeetingType"), 1, 18, forceRebuild
    EnsureTable ws, T_CONFIG, Array("Key", "Value"), 1, 20, forceRebuild
End Sub

Private Sub EnsureLedgerTable(ByVal forceRebuild As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_LEDGER)

    Dim headers As Variant
    headers = Array( _
        "TxnID", "Date", "MonthKey", "FiscalYear", _
        "TxnType", "TxnSubtype", _
        "Category", "Event", "Charity", _
        "Gross", "Fees", "Net", _
        "PaymentMethod", "SourceType", "SourceName", _
        "MemberID", "MemberName", "MemberEmail", "Memo", _
        "ReceiptRequired", "ReceiptStatus", "ReceiptInfoID", _
        "ApprovedBy", "ClosedFlag", _
        "ExternalSource", "ExternalTxnID", "ImportBatchID", _
        "CreatedAt", "UpdatedAt" _
    )

    EnsureTable ws, T_LEDGER, headers, 1, 1, forceRebuild
End Sub

Private Sub EnsureReceiptsTable(ByVal forceRebuild As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_RECEIPTS)

    EnsureTable ws, T_RECEIPTS, Array( _
        "ReceiptInfoID", "TxnID", "ReceiptNumber", "Vendor", "ReceiptDate", "ReceivedDate", _
        "StorageLocation", "Notes", "WaivedReason", "WaivedBy", "WaivedAt", "VerifiedFlag" _
    ), 1, 1, forceRebuild
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

Private Sub EnsureMembersTable(ByVal forceRebuild As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_MEMBERS)
    EnsureTable ws, T_MEMBERS, Array( _
        "MemberID", "MemberName", "MemberEmail", "MembershipType", "Status", _
        "ApplicationReceivedDate", "BoardVoteDate", "BoardVoteResult", "ApprovalNotes", _
        "JoinedDate", "DuesPaidFlag", "DuesPaidDate", "DuesAmount", _
        "DuesSplitPlanFlag", "DuesSecondInstallmentDue", "RenewalDate", _
        "RenewalIntervalMonths", "ParticipationYear", "EventsAttendedYTD", "VolunteerEventsYTD", _
        "AgeVerifiedFlag", "AgeVerifiedBy", "AgeVerifiedAt", "Notes", _
        "ExternalSource", "ExternalMemberID", "LastUpdatedAt" _
    ), 1, 1, forceRebuild
End Sub

Private Sub EnsureMeetingsTables(ByVal forceRebuild As Boolean)
    EnsureTable GetSheet(SH_MEETINGS), T_MEETINGS, Array( _
        "MeetingID", "MeetingDate", "StartTime", "EndTime", "Scribe", "Location", _
        "MeetingType", "OpenSessionFlag", "RemoteAllowedFlag", "RemoteRequestBy", _
        "RemoteRequestAt", "NoticeSentAt", "QuorumMetFlag", _
        "MinutesDocPath", "MinutesPdfPath", "CreatedAt" _
    ), 1, 1, forceRebuild

    RenameTableOnSheet SH_ATTENDANCE, "tblAttendance", T_ATTENDANCE
    EnsureTable GetSheet(SH_ATTENDANCE), T_ATTENDANCE, Array( _
        "MeetingID", "PersonName", "Role", "PresentFlag" _
    ), 1, 1, forceRebuild

    EnsureTable GetSheet(SH_MINUTES_LINES), T_MINUTES_LINES, Array( _
        "MeetingID", "LineTime", "Topic", "Notes", "ActionItem", "Owner", "DueDate", "ActionStatus" _
    ), 1, 1, forceRebuild

    EnsureTable GetSheet(SH_ACTION_ITEMS), T_ACTION_ITEMS, Array( _
        "ActionID", "MeetingID", "ActionItem", "Owner", "DueDate", "Status", "Notes" _
    ), 1, 1, forceRebuild
End Sub

Private Sub EnsureAgendaTable(ByVal forceRebuild As Boolean)
    EnsureTable GetSheet(SH_AGENDA), T_AGENDA, Array( _
        "AgendaID", "AgendaDate", "DocPath", "PdfPath", "CreatedAt" _
    ), 1, 1, forceRebuild
End Sub

Private Sub EnsureComplianceTables(ByVal forceRebuild As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_COMPLIANCE)

    EnsureTable ws, T_COMPLIANCE_EVENTS, Array( _
        "EventID", "EventName", "EventDate", "StartTime", "EndTime", "Location", "VenueContact", _
        "EventType", "AgeRestriction", "CharityPartner", "Purpose", "SafetyLead", "FinanceLead", _
        "DoorCheckinLead", "VolunteerLead", "PaymentMethodsOnSite", "CashHandlingPlan", _
        "IncidentReportingPostedFlag", "PhotoMediaConsentPlan", "ApprovalStatus", "ApprovedBy", _
        "ApprovedAt", "Notes", "CreatedAt", "UpdatedAt" _
    ), 1, 1, forceRebuild

    EnsureTable ws, T_EVENT_ATTENDANCE, Array( _
        "EventID", "MemberID", "MemberName", "AttendedFlag", "VolunteerFlag", "VolunteerRole", _
        "Hours", "Notes" _
    ), 40, 1, forceRebuild

    EnsureTable ws, T_COMPLIANCE_VENDORS, Array( _
        "VendorID", "BusinessName", "ContactName", "ContactEmail", "ContactPhone", _
        "ServicesProvided", "AgreementPath", "InsurancePath", "PaymentTerms", _
        "ApprovedFlag", "ApprovedBy", "ApprovedAt", "Notes", "CreatedAt", "UpdatedAt" _
    ), 80, 1, forceRebuild

    EnsureTable ws, T_COMPLIANCE_VOLUNTEERS, Array( _
        "VolunteerID", "EventID", "VolunteerName", "VolunteerEmail", "VolunteerPhone", _
        "EmergencyContact", "AgreementSignedFlag", "SignedAt", "AgreementPath", "Notes", _
        "CreatedAt", "UpdatedAt" _
    ), 120, 1, forceRebuild

    EnsureTable ws, T_COMPLIANCE_GIVEAWAYS, Array( _
        "GiveawayID", "Title", "Sponsor", "StartDate", "EndDate", "PrizeDescription", _
        "PrizeValue", "NoPurchaseRequiredFlag", "FreeEntryMethod", "Eligibility", _
        "WinnerSelectionMethod", "WinnerNotifiedAt", "WinnerName", "WinnerContact", _
        "RulesDocPath", "Notes", "CreatedAt", "UpdatedAt" _
    ), 170, 1, forceRebuild

    EnsureTable ws, T_COMPLIANCE_REFUNDS, Array( _
        "RefundID", "RequestDate", "RequesterName", "RequesterEmail", "TxnID", "EventID", _
        "AmountRequested", "Reason", "Status", "ApprovedBy", "ApprovedAt", "PaidTxnID", _
        "FormDocPath", "Notes", "CreatedAt", "UpdatedAt" _
    ), 230, 1, forceRebuild

    EnsureTable ws, T_COMPLIANCE_INCIDENTS, Array( _
        "CaseID", "CaseType", "ReportDate", "ReportedByName", "ReportedByContact", _
        "EventID", "Location", "InvolvedParties", "Summary", "ImmediateActionsTaken", _
        "Severity", "SSC_ConsentRelatedFlag", "ConfidentialFlag", "Status", "OutcomeID", _
        "RetentionUntil", "EvidencePath", "Notes", "CreatedAt", "UpdatedAt" _
    ), 300, 1, forceRebuild

    EnsureTable ws, T_COMPLIANCE_OUTCOMES, Array( _
        "OutcomeID", "CaseID", "OutcomeDate", "Findings", "ConsequenceType", _
        "ConsequenceDetails", "EffectiveFrom", "EffectiveTo", "NoticeDocPath", _
        "AppealAllowedFlag", "AppealDeadline", "ClosedAt", "CreatedAt" _
    ), 370, 1, forceRebuild

    EnsureTable ws, T_COMPLIANCE_APPEALS, Array( _
        "AppealID", "OutcomeID", "AppealDate", "AppellantName", "Grounds", _
        "EvidencePath", "Status", "DecisionDate", "DecisionBy", "DecisionNotes", _
        "AppealDocPath", "CreatedAt" _
    ), 430, 1, forceRebuild

    EnsureTable ws, T_COMPLIANCE_ACKS, Array( _
        "AckID", "PersonName", "PersonEmail", "PolicyName", "PolicyVersion", _
        "SignedFlag", "SignedAt", "ProofPath", "Notes" _
    ), 480, 1, forceRebuild

    EnsureTable ws, T_COMPLIANCE_PAYMENT_ADMINS, Array( _
        "AdminID", "Platform", "AccountIdentifier", "AdminName", "AdminEmail", _
        "Role", "TwoFAEnabledFlag", "AddedAt", "RemovedAt", "LastReviewedAt", _
        "ReviewedInMeetingID", "Notes" _
    ), 520, 1, forceRebuild
End Sub

Private Sub EnsureImportsTables(ByVal forceRebuild As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_IMPORTS)

    EnsureTable ws, T_IMPORTLOG, Array( _
        "ImportBatchID", "Source", "ImportedAt", "FileName", "FileHash", "RowCount", "Notes", "Status" _
    ), 1, 1, forceRebuild
    EnsureTable ws, T_ZEFFY_RAW, Array( _
        "ImportBatchID", "RowNum", "RawLine", "ExternalTxnID", "RowHash", "CreatedAt" _
    ), 1, 9, forceRebuild
    EnsureTable ws, T_BLAZE_RAW, Array( _
        "ImportBatchID", "RowNum", "RawLine", "ExternalTxnID", "RowHash", "CreatedAt" _
    ), 1, 17, forceRebuild
    EnsureTable ws, T_IMPORTMAP_ZEFFY, Array( _
        "SourceColumn", "TargetTable", "TargetColumn", "TransformRule", "ActiveFlag" _
    ), 1, 25, forceRebuild
    EnsureTable ws, T_IMPORTMAP_BLAZE, Array( _
        "SourceColumn", "TargetTable", "TargetColumn", "TransformRule", "ActiveFlag" _
    ), 1, 31, forceRebuild
End Sub

Private Sub EnsureAuditTable(ByVal forceRebuild As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_AUDIT)
    EnsureTable ws, T_AUDIT, Array("Timestamp", "User", "Action", "TxnID", "Details"), 1, 1, forceRebuild
End Sub

Private Sub EnsureErrorLogTable(ByVal forceRebuild As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_ERRORLOG)
    EnsureTable ws, T_ERRORLOG, Array( _
        "ErrorID", "Timestamp", "User", "Procedure", "ErrNumber", "ErrDescription", "Context", "Stack" _
    ), 1, 1, forceRebuild
End Sub

Private Sub EnsureArchiveTable(ByVal forceRebuild As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_ARCHIVE)
    EnsureTable ws, "ARCH_BoardPackets", Array( _
        "MonthKey", "FiscalYear", "GeneratedAt", "GeneratedBy", "PdfPath", "SnapshotRange", "SnapshotHash" _
    ), 1, 1, forceRebuild
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
    ws.Range("G17").value = "Raised (YTD)"
    ws.Range("G18").value = "Paid Out (YTD)"

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

Private Sub EnsureComplianceReportLayout()
    Dim ws As Worksheet
    Set ws = GetSheet(SH_COMPLIANCE_REPORT)
    ws.Cells.ClearContents
    ws.Range("A1").value = "TCPP Compliance Export"
    ws.Range("A2").value = "Generated"
    ws.Range("B2").value = Now
    ws.Columns("A:B").ColumnWidth = 28
End Sub

Private Sub EnsureConfigDefaults()
    Dim cfg As ListObject
    Set cfg = GetTable(SH_LOOKUPS, T_CONFIG)

    UpsertConfig cfg, CFG_FISCAL_START_MONTH, "6"
    UpsertConfig cfg, CFG_APPROVER_NAME, Application.UserName
    UpsertConfig cfg, CFG_RENEWAL_INTERVAL, "12"
    UpsertConfig cfg, CFG_STRICT_BUDGET, "FALSE"
    UpsertConfig cfg, CFG_RECEIPT_WAIVER_REASON, "TRUE"

    UpsertConfig cfg, CFG_PATH_BOARDPACKETS, ".\BoardPackets\"
    UpsertConfig cfg, CFG_PATH_MINUTES_DOCX, ".\Minutes\DOCX\"
    UpsertConfig cfg, CFG_PATH_MINUTES_PDF, ".\Minutes\PDF\"
    UpsertConfig cfg, CFG_PATH_AGENDA_DOCX, ".\Agenda\DOCX\"
    UpsertConfig cfg, CFG_PATH_AGENDA_PDF, ".\Agenda\PDF\"
    UpsertConfig cfg, CFG_PATH_IMPORTS_ZEFFY, ".\Imports\Zeffy\"
    UpsertConfig cfg, CFG_PATH_IMPORTS_BLAZE, ".\Imports\Blaze\"
    UpsertConfig cfg, CFG_PATH_COMPLIANCE_DOCS, ".\Compliance\Docs\"
    UpsertConfig cfg, CFG_PATH_IMPORTS_ARCHIVE, ".\Imports\Archive\"

    UpsertConfig cfg, CFG_MEETING_NOTICE_REGULAR, "14"
    UpsertConfig cfg, CFG_MEETING_NOTICE_ANNUAL, "30"
    UpsertConfig cfg, CFG_MEETING_NOTICE_SPECIAL, "48"
    UpsertConfig cfg, CFG_BOARD_QUORUM, "5"
    UpsertConfig cfg, CFG_DUES_INITIAL_DUE_DAYS, "30"
    UpsertConfig cfg, CFG_DUES_AMOUNT_FULL, "30"
    UpsertConfig cfg, CFG_DUES_AMOUNT_ATLARGE, "20"
    UpsertConfig cfg, CFG_DUES_SPLIT_FULL, "TRUE"
    UpsertConfig cfg, CFG_DUES_SPLIT_ATLARGE, "FALSE"
    UpsertConfig cfg, CFG_MIN_AGE_FULL, "21"
    UpsertConfig cfg, CFG_MIN_AGE_ATLARGE, "18"
    UpsertConfig cfg, CFG_EVENTS_REQ_FULL, "4"
    UpsertConfig cfg, CFG_VOLUNTEER_REQ_FULL, "1"
    UpsertConfig cfg, CFG_RETENTION_MINOR, "3"
    UpsertConfig cfg, CFG_RETENTION_SEVERE, "7"
    UpsertConfig cfg, CFG_REFUND_APPROVAL_SELF_OK, "FALSE"
End Sub

Private Sub EnsureDefaultFolders()
    EnsureFolderPath ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_BOARDPACKETS, ".\BoardPackets\"))
    EnsureFolderPath ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_MINUTES_DOCX, ".\Minutes\DOCX\"))
    EnsureFolderPath ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_MINUTES_PDF, ".\Minutes\PDF\"))
    EnsureFolderPath ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_AGENDA_DOCX, ".\Agenda\DOCX\"))
    EnsureFolderPath ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_AGENDA_PDF, ".\Agenda\PDF\"))
    EnsureFolderPath ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_IMPORTS_ZEFFY, ".\Imports\Zeffy\"))
    EnsureFolderPath ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_IMPORTS_BLAZE, ".\Imports\Blaze\"))
    EnsureFolderPath ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_COMPLIANCE_DOCS, ".\Compliance\Docs\"))
    EnsureFolderPath ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_IMPORTS_ARCHIVE, ".\Imports\Archive\"))
End Sub

Private Sub SeedLookupsIfEmpty()
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

    Dim types As ListObject: Set types = GetTable(SH_LOOKUPS, T_TXN_TYPES)
    If types.DataBodyRange Is Nothing Then
        AppendListValue types, 1, "Income"
        AppendListValue types, 1, "Expense"
        AppendListValue types, 1, "Reimbursement"
        AppendListValue types, 1, "Transfer"
        AppendListValue types, 1, "Deposit"
        AppendListValue types, 1, "Withdrawal"
        AppendListValue types, 1, "Adjustment"
    End If

    Dim subtypes As ListObject: Set subtypes = GetTable(SH_LOOKUPS, T_TXN_SUBTYPES)
    If subtypes.DataBodyRange Is Nothing Then
        AppendListValue subtypes, 1, "Dues"
        AppendListValue subtypes, 1, "Donation"
        AppendListValue subtypes, 1, "Event"
        AppendListValue subtypes, 1, "Operations"
        AppendListValue subtypes, 1, "Labor"
        AppendListValue subtypes, 1, "Insurance"
        AppendListValue subtypes, 1, "Rent"
        AppendListValue subtypes, 1, "Grant"
    End If

    Dim ev As ListObject: Set ev = GetTable(SH_LOOKUPS, T_EVENTS_LIST)
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

    Dim mt As ListObject: Set mt = GetTable(SH_LOOKUPS, T_MEMBERTYPES)
    If mt.DataBodyRange Is Nothing Then
        AppendListValue mt, 1, "Full"
        AppendListValue mt, 1, "AtLarge"
        AppendListValue mt, 1, "Other"
    End If

    Dim mtg As ListObject: Set mtg = GetTable(SH_LOOKUPS, T_MEETING_TYPES)
    If mtg.DataBodyRange Is Nothing Then
        AppendListValue mtg, 1, "Regular"
        AppendListValue mtg, 1, "Annual"
        AppendListValue mtg, 1, "Special"
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
' Error handling
'========================

Public Sub HandleError(ByVal procName As String, ByVal errObj As ErrObject, ByVal context As String)
    LogError procName, errObj, context
    MsgBox "Something went wrong. Details were logged." & vbCrLf & errObj.Description, vbExclamation, "TCPP"
End Sub

Public Sub LogError(ByVal procName As String, ByVal errObj As ErrObject, ByVal context As String)
    On Error Resume Next
    Dim lo As ListObject: Set lo = GetTable(SH_ERRORLOG, T_ERRORLOG)
    Dim lr As ListRow: Set lr = lo.ListRows.Add

    lr.Range.Cells(1, 1).value = "ERR-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(Int((9999 * Rnd) + 1), "0000")
    lr.Range.Cells(1, 2).value = Now
    lr.Range.Cells(1, 3).value = Application.UserName
    lr.Range.Cells(1, 4).value = procName
    lr.Range.Cells(1, 5).value = errObj.Number
    lr.Range.Cells(1, 6).value = errObj.Description
    lr.Range.Cells(1, 7).value = context
    lr.Range.Cells(1, 8).value = ""
End Sub

'========================
' UI theming helpers
'========================

Public Sub ApplyTheme(ByVal frm As Object)
    On Error Resume Next
    frm.BackColor = RGB(245, 246, 250)
    frm.Font.name = "Segoe UI"
    frm.Font.Size = 9

    Dim ctl As Object
    For Each ctl In frm.Controls
        ctl.Font.name = "Segoe UI"
        ctl.Font.Size = 9
        Select Case TypeName(ctl)
            Case "CommandButton"
                ctl.BackColor = RGB(32, 85, 154)
                ctl.ForeColor = RGB(255, 255, 255)
            Case "Label"
                ctl.BackStyle = 0
                ctl.ForeColor = RGB(50, 50, 50)
            Case "TextBox", "ComboBox", "ListBox"
                ctl.BackColor = RGB(255, 255, 255)
                ctl.ForeColor = RGB(30, 30, 30)
        End Select
    Next ctl
End Sub

'========================
' Sheet/table helpers
'========================

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
End Function

Private Sub RenameLookupTable(ByVal oldName As String, ByVal newName As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet(SH_LOOKUPS)
    If ws Is Nothing Then Exit Sub
    Dim lo As ListObject: Set lo = ws.ListObjects(oldName)
    If lo Is Nothing Then Exit Sub
    lo.name = newName
    On Error GoTo 0
End Sub

Private Sub RenameTableOnSheet(ByVal sheetName As String, ByVal oldName As String, ByVal newName As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet(sheetName)
    If ws Is Nothing Then Exit Sub
    Dim lo As ListObject: Set lo = ws.ListObjects(oldName)
    If lo Is Nothing Then Exit Sub
    lo.name = newName
    On Error GoTo 0
End Sub

Private Function FindTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set FindTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTable Is Nothing Then Exit Function
    Next ws
End Function

Public Sub EnsureColumns(ByVal tableName As String, ByVal requiredHeaders As Variant)
    Dim lo As ListObject
    Set lo = FindTable(tableName)
    If lo Is Nothing Then Exit Sub
    EnsureHeaders lo, requiredHeaders
End Sub

Public Sub EnsureFolders(ByVal relativePaths As Variant)
    Dim i As Long
    For i = LBound(relativePaths) To UBound(relativePaths)
        EnsureFolderPath ResolveWorkbookRelativePath(CStr(relativePaths(i)))
    Next i
End Sub

Public Function SafeWriteRow(ByVal lo As ListObject, ByVal monthKey As String) As ListRow
    If lo Is Nothing Then Exit Function
    If lo.name = T_LEDGER And Len(monthKey) > 0 Then
        If IsMonthClosed(monthKey) Then Err.Raise vbObjectError + 900, "SafeWriteRow", "Month is closed"
    End If
    Set SafeWriteRow = lo.ListRows.Add
End Function

Public Sub SafeUpdateRow(ByVal lo As ListObject, ByVal monthKey As String)
    If lo Is Nothing Then Exit Sub
    If lo.name = T_LEDGER And Len(monthKey) > 0 Then
        If IsMonthClosed(monthKey) Then Err.Raise vbObjectError + 901, "SafeUpdateRow", "Month is closed"
    End If
End Sub

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

Public Function GetOrCreateSheet(ByVal sheetName As String, ByVal veryHidden As Boolean) As Worksheet
    EnsureSheet sheetName, IIf(veryHidden, xlSheetVeryHidden, xlSheetVisible)
    Set GetOrCreateSheet = GetSheet(sheetName)
End Function

Public Function GetOrCreateTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal headers As Variant) As ListObject
    EnsureTable ws, tableName, headers, 1, 1, False
    Set GetOrCreateTable = ws.ListObjects(tableName)
End Function

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

Private Function ColumnExists(ByVal lo As ListObject, ByVal columnName As String) As Boolean
    On Error Resume Next
    ColumnExists = Not lo.ListColumns(columnName) Is Nothing
    On Error GoTo 0
End Function

Private Sub SetIfExists(ByVal lo As ListObject, ByVal lr As ListRow, ByVal columnName As String, ByVal value As Variant)
    If lo Is Nothing Then Exit Sub
    If lr Is Nothing Then Exit Sub
    If Not ColumnExists(lo, columnName) Then Exit Sub
    lr.Range.Cells(1, lo.ListColumns(columnName).Index).value = value
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

Public Function GetConfigValue(ByVal key As String, Optional ByVal defaultValue As String = "") As String
    On Error GoTo EH
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
    Exit Function
EH:
    HandleError "GetConfigValue", Err, key
    GetConfigValue = defaultValue
End Function

'========================
' Fiscal/month helpers
'========================

Public Function FiscalYearForMonthKey(ByVal monthKey As String) As Long
    On Error GoTo EH
    Dim y As Long, m As Long
    y = CLng(left$(monthKey, 4))
    m = CLng(right$(monthKey, 2))

    Dim startM As Long
    startM = CLng(GetConfigValue(CFG_FISCAL_START_MONTH, "6"))

    If m >= startM Then
        FiscalYearForMonthKey = y + 1
    Else
        FiscalYearForMonthKey = y
    End If
    Exit Function
EH:
    HandleError "FiscalYearForMonthKey", Err, monthKey
    FiscalYearForMonthKey = 0
End Function

Public Function MonthKeyFromDate(ByVal d As Date) As String
    On Error GoTo EH
    MonthKeyFromDate = Format(d, "yyyy-mm")
    Exit Function
EH:
    HandleError "MonthKeyFromDate", Err, ""
    MonthKeyFromDate = ""
End Function

Public Function MonthKeyForDate(ByVal d As Date) As String
    MonthKeyForDate = MonthKeyFromDate(d)
End Function

Public Function FiscalYearForDate(ByVal d As Date) As Long
    FiscalYearForDate = FiscalYearForMonthKey(MonthKeyFromDate(d))
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
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_AUDIT, T_AUDIT)
    Dim lr As ListRow: Set lr = lo.ListRows.Add
    lr.Range.Cells(1, 1).value = Now
    lr.Range.Cells(1, 2).value = Application.UserName
    lr.Range.Cells(1, 3).value = action
    lr.Range.Cells(1, 4).value = txnId
    lr.Range.Cells(1, 5).value = details
    Exit Sub
EH:
    HandleError "AuditLog", Err, action
End Sub

'========================
' Ledger operations
'========================

Public Function NextTxnId(ByVal monthKey As String) As String
    On Error GoTo EH
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
    Exit Function
EH:
    HandleError "NextTxnId", Err, monthKey
    NextTxnId = ""
End Function

Public Function AddLedgerEntry( _
    ByVal txnDate As Date, ByVal txnType As String, ByVal txnSubtype As String, _
    ByVal category As String, ByVal eventName As String, ByVal charityName As String, _
    ByVal gross As Double, ByVal fees As Double, ByVal paymentMethod As String, _
    ByVal sourceType As String, ByVal sourceName As String, _
    ByVal memberName As String, ByVal memberEmail As String, _
    ByVal memo As String, ByVal receiptRequired As Boolean, _
    Optional ByVal externalSource As String = "Manual", Optional ByVal externalTxnId As String = "", _
    Optional ByVal importBatchId As String = "", Optional ByVal allowInClosedMonth As Boolean = False _
) As String

    On Error GoTo EH
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
    lr.Range.Cells(1, lo.ListColumns("TxnSubtype").Index).value = txnSubtype

    lr.Range.Cells(1, lo.ListColumns("Category").Index).value = category
    lr.Range.Cells(1, lo.ListColumns("Event").Index).value = eventName
    lr.Range.Cells(1, lo.ListColumns("Charity").Index).value = charityName

    lr.Range.Cells(1, lo.ListColumns("Gross").Index).value = gross
    lr.Range.Cells(1, lo.ListColumns("Fees").Index).value = fees
    lr.Range.Cells(1, lo.ListColumns("Net").Index).value = gross - fees

    lr.Range.Cells(1, lo.ListColumns("PaymentMethod").Index).value = paymentMethod
    lr.Range.Cells(1, lo.ListColumns("SourceType").Index).value = sourceType
    lr.Range.Cells(1, lo.ListColumns("SourceName").Index).value = sourceName

    If ColumnExists(lo, "MemberID") Then
        lr.Range.Cells(1, lo.ListColumns("MemberID").Index).value = FindMemberIdByEmail(memberEmail)
    End If
    lr.Range.Cells(1, lo.ListColumns("MemberName").Index).value = memberName
    lr.Range.Cells(1, lo.ListColumns("MemberEmail").Index).value = memberEmail
    lr.Range.Cells(1, lo.ListColumns("Memo").Index).value = memo

    If (LCase$(txnType) = "expense" Or LCase$(txnType) = "reimbursement") And receiptRequired = False Then
        receiptRequired = True
    End If

    lr.Range.Cells(1, lo.ListColumns("ReceiptRequired").Index).value = receiptRequired
    lr.Range.Cells(1, lo.ListColumns("ReceiptStatus").Index).value = IIf(receiptRequired, "Missing", "NotRequired")
    lr.Range.Cells(1, lo.ListColumns("ReceiptInfoID").Index).value = ""

    lr.Range.Cells(1, lo.ListColumns("ApprovedBy").Index).value = GetConfigValue(CFG_APPROVER_NAME, Application.UserName)
    lr.Range.Cells(1, lo.ListColumns("ClosedFlag").Index).value = False

    lr.Range.Cells(1, lo.ListColumns("ExternalSource").Index).value = externalSource
    lr.Range.Cells(1, lo.ListColumns("ExternalTxnID").Index).value = externalTxnId
    lr.Range.Cells(1, lo.ListColumns("ImportBatchID").Index).value = importBatchId

    lr.Range.Cells(1, lo.ListColumns("CreatedAt").Index).value = Now
    lr.Range.Cells(1, lo.ListColumns("UpdatedAt").Index).value = Now

    AuditLog "Create", txnId, txnType & " / " & monthKey & " / " & Format$(gross - fees, "0.00")

    AddLedgerEntry = txnId
    Exit Function
EH:
    HandleError "AddLedgerEntry", Err, MonthKeyFromDate(txnDate)
    AddLedgerEntry = ""
End Function

Public Sub UpdateLedgerFields(ByVal txnId As String, ByVal category As String, ByVal eventName As String, ByVal charityName As String, _
                             ByVal receiptRequired As Boolean)
    On Error GoTo EH
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
    Dim infoCol As Long: infoCol = lo.ListColumns("ReceiptInfoID").Index

    Dim curInfo As String: curInfo = NzStr(lo.DataBodyRange.Cells(rowIdx, infoCol).value, "")
    Dim curStatus As String: curStatus = NzStr(lo.DataBodyRange.Cells(rowIdx, statusCol).value, "")

    If receiptRequired Then
        If Len(curInfo) = 0 Then
            If curStatus <> "Waived" Then lo.DataBodyRange.Cells(rowIdx, statusCol).value = "Missing"
        Else
            lo.DataBodyRange.Cells(rowIdx, statusCol).value = "Recorded"
        End If
    Else
        lo.DataBodyRange.Cells(rowIdx, statusCol).value = "NotRequired"
    End If

    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("UpdatedAt").Index).value = Now
    AuditLog "Edit", txnId, "Category/Event/Charity/ReceiptRequired"
    Exit Sub
EH:
    HandleError "UpdateLedgerFields", Err, txnId
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

Public Function CreateReceiptInfo(ByVal txnId As String, ByVal vendor As String, ByVal receiptDate As Variant, _
                                  ByVal receivedDate As Date, ByVal storageLocation As String, ByVal notes As String) As String
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_RECEIPTS, T_RECEIPTS)
    Dim lr As ListRow: Set lr = lo.ListRows.Add

    Dim receiptId As String
    receiptId = "RCPT-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(Int((9999 * Rnd) + 1), "0000")

    lr.Range.Cells(1, lo.ListColumns("ReceiptInfoID").Index).value = receiptId
    lr.Range.Cells(1, lo.ListColumns("TxnID").Index).value = txnId
    lr.Range.Cells(1, lo.ListColumns("ReceiptNumber").Index).value = ""
    lr.Range.Cells(1, lo.ListColumns("Vendor").Index).value = vendor
    If IsDate(receiptDate) Then lr.Range.Cells(1, lo.ListColumns("ReceiptDate").Index).value = CDate(receiptDate)
    lr.Range.Cells(1, lo.ListColumns("ReceivedDate").Index).value = receivedDate
    lr.Range.Cells(1, lo.ListColumns("StorageLocation").Index).value = storageLocation
    lr.Range.Cells(1, lo.ListColumns("Notes").Index).value = notes
    lr.Range.Cells(1, lo.ListColumns("VerifiedFlag").Index).value = True

    Dim lLedger As ListObject: Set lLedger = GetTable(SH_LEDGER, T_LEDGER)
    Dim rowIdx As Long: rowIdx = FindLedgerRowIndex(txnId)
    If rowIdx > 0 Then
        lLedger.DataBodyRange.Cells(rowIdx, lLedger.ListColumns("ReceiptInfoID").Index).value = receiptId
        lLedger.DataBodyRange.Cells(rowIdx, lLedger.ListColumns("ReceiptStatus").Index).value = "Recorded"
        lLedger.DataBodyRange.Cells(rowIdx, lLedger.ListColumns("UpdatedAt").Index).value = Now
    End If

    AuditLog "ReceiptRecorded", txnId, receiptId
    CreateReceiptInfo = receiptId
    Exit Function
EH:
    HandleError "CreateReceiptInfo", Err, txnId
    CreateReceiptInfo = ""
End Function

Public Sub WaiveReceipt(ByVal txnId As String, ByVal reason As String)
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    Dim rowIdx As Long: rowIdx = FindLedgerRowIndex(txnId)
    If rowIdx = 0 Then Err.Raise vbObjectError + 518, "WaiveReceipt", "TxnID not found: " & txnId

    If Len(Trim$(reason)) = 0 Then
        Err.Raise vbObjectError + 518, "WaiveReceipt", "Waive reason is required."
    End If

    Dim monthKey As String: monthKey = CStr(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("MonthKey").Index).value)
    If IsMonthClosed(monthKey) Then Err.Raise vbObjectError + 519, "WaiveReceipt", "Month is closed: " & monthKey

    Dim receiptId As String
    receiptId = CreateReceiptInfo(txnId, "", Empty, Date, "", "")

    Dim rLo As ListObject: Set rLo = GetTable(SH_RECEIPTS, T_RECEIPTS)
    Dim rIdx As Long: rIdx = FindReceiptRowIndex(receiptId)

    If rIdx > 0 Then
        rLo.DataBodyRange.Cells(rIdx, rLo.ListColumns("WaivedReason").Index).value = reason
        rLo.DataBodyRange.Cells(rIdx, rLo.ListColumns("WaivedBy").Index).value = GetConfigValue(CFG_APPROVER_NAME, Application.UserName)
        rLo.DataBodyRange.Cells(rIdx, rLo.ListColumns("WaivedAt").Index).value = Now
        rLo.DataBodyRange.Cells(rIdx, rLo.ListColumns("VerifiedFlag").Index).value = False
    End If

    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ReceiptStatus").Index).value = "Waived"
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("UpdatedAt").Index).value = Now

    AuditLog "WaiveReceipt", txnId, reason
    Exit Sub
EH:
    HandleError "WaiveReceipt", Err, txnId
End Sub

Private Function FindReceiptRowIndex(ByVal receiptId As String) As Long
    Dim lo As ListObject: Set lo = GetTable(SH_RECEIPTS, T_RECEIPTS)
    FindReceiptRowIndex = 0
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Range
    For Each r In lo.ListColumns("ReceiptInfoID").DataBodyRange.Cells
        If CStr(r.value) = receiptId Then
            FindReceiptRowIndex = r.Row - lo.DataBodyRange.Row + 1
            Exit Function
        End If
    Next r
End Function

'========================
' Month status / reconcile / close
'========================

Public Function IsMonthClosed(ByVal monthKey As String) As Boolean
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_MONTHSTATUS, T_MONTHSTATUS)
    Dim rowIdx As Long: rowIdx = FindMonthStatusRow(monthKey)
    If rowIdx = 0 Then
        IsMonthClosed = False
    Else
        IsMonthClosed = CBool(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ClosedFlag").Index).value)
    End If
    Exit Function
EH:
    HandleError "IsMonthClosed", Err, monthKey
    IsMonthClosed = False
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
    On Error GoTo EH
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

            If LCase$(t) = "income" Or LCase$(t) = "deposit" Then
                deposits = deposits + n
            ElseIf LCase$(t) = "expense" Or LCase$(t) = "reimbursement" Or LCase$(t) = "withdrawal" Then
                withdrawals = withdrawals + Abs(n)
            ElseIf LCase$(t) = "adjustment" Then
                If n >= 0 Then deposits = deposits + n Else withdrawals = withdrawals + Abs(n)
            End If
        End If
    Next i
    Exit Sub
EH:
    HandleError "ComputeMonthLedgerTotals", Err, monthKey
End Sub

Public Sub SaveReconciliation(ByVal monthKey As String, ByVal beginningBal As Double, ByVal endingBal As Double)
    On Error GoTo EH
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
    Exit Sub
EH:
    HandleError "SaveReconciliation", Err, monthKey
End Sub

Public Function GateCheckMonth(ByVal monthKey As String, ByRef missingCategories As Long, ByRef missingReceipts As Long, _
                               ByRef missingReceiptsAmt As Double, ByRef reconOk As Boolean, _
                               ByRef charityImbalance As Boolean, ByRef budgetOverrun As Boolean) As String
    On Error GoTo EH
    missingCategories = 0
    missingReceipts = 0
    missingReceiptsAmt = 0#
    reconOk = False
    charityImbalance = False
    budgetOverrun = False

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
                    If rs <> "Recorded" And rs <> "Waived" Then
                        missingReceipts = missingReceipts + 1
                        missingReceiptsAmt = missingReceiptsAmt + Abs(CDbl(lo.DataBodyRange.Cells(i, netCol).value))
                    End If
                End If
            End If
        Next i
    End If

    reconOk = IsReconOk(monthKey)
    charityImbalance = (CharityHeldYTD(monthKey) < 0#) Or CharityPaidExceedsRaisedYTD(monthKey)
    budgetOverrun = HasBudgetOverrun(monthKey)

    Dim msg As String
    msg = "Month " & monthKey & " gates:" & vbCrLf & _
          "Uncategorized: " & CStr(missingCategories) & vbCrLf & _
          "Missing receipts: " & CStr(missingReceipts) & " ($" & Format$(missingReceiptsAmt, "0.00") & ")" & vbCrLf & _
          "Reconciled: " & IIf(reconOk, "YES", "NO") & vbCrLf & _
          "Charity imbalance: " & IIf(charityImbalance, "YES", "NO") & vbCrLf & _
          "Budget overrun: " & IIf(budgetOverrun, "YES", "NO")
    GateCheckMonth = msg
    Exit Function
EH:
    HandleError "GateCheckMonth", Err, monthKey
    GateCheckMonth = "Gate check failed."
End Function

Public Function IsReconOk(ByVal monthKey As String) As Boolean
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_MONTHSTATUS, T_MONTHSTATUS)
    Dim rowIdx As Long: rowIdx = FindMonthStatusRow(monthKey)
    If rowIdx = 0 Then
        IsReconOk = False
        Exit Function
    End If

    Dim status As String: status = CStr(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ReconStatus").Index).value)
    IsReconOk = (status = "OK")
    Exit Function
EH:
    HandleError "IsReconOk", Err, monthKey
    IsReconOk = False
End Function

Public Sub CloseMonth(ByVal monthKey As String)
    Dim prevSU As Boolean, prevEV As Boolean
    prevSU = Application.ScreenUpdating
    prevEV = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo EH
    Dim mc As Long, mr As Long
    Dim mAmt As Double, reconOk As Boolean, charityImbalance As Boolean, budgetOverrun As Boolean
    Dim gateMsg As String
    gateMsg = GateCheckMonth(monthKey, mc, mr, mAmt, reconOk, charityImbalance, budgetOverrun)

    Dim strictBudget As Boolean: strictBudget = (UCase$(GetConfigValue(CFG_STRICT_BUDGET, "FALSE")) = "TRUE")

    If mc <> 0 Or mr <> 0 Or (Not reconOk) Or charityImbalance Or (strictBudget And budgetOverrun) Then
        Err.Raise vbObjectError + 520, "CloseMonth", "Close blocked." & vbCrLf & gateMsg
    End If

    EnsureMonthStatusRow monthKey

    Dim ms As ListObject: Set ms = GetTable(SH_MONTHSTATUS, T_MONTHSTATUS)
    Dim rowIdx As Long: rowIdx = FindMonthStatusRow(monthKey)
    ms.DataBodyRange.Cells(rowIdx, ms.ListColumns("ClosedFlag").Index).value = True
    ms.DataBodyRange.Cells(rowIdx, ms.ListColumns("ClosedAt").Index).value = Now

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
    GoTo CleanExit
EH:
    HandleError "CloseMonth", Err, monthKey
CleanExit:
    Application.EnableEvents = prevEV
    Application.ScreenUpdating = prevSU
End Sub

Private Sub ProtectDataSheets()
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
    Dim prevSU As Boolean, prevEV As Boolean
    prevSU = Application.ScreenUpdating
    prevEV = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo EH

    Dim ws As Worksheet: Set ws = GetSheet(SH_REPORT)
    Dim oldVis As XlSheetVisibility: oldVis = ws.Visible
    ws.Visible = xlSheetVisible

    EnsureReportSheetLayout

    Dim fy As Long: fy = FiscalYearForMonthKey(monthKey)
    ws.Range("B3").value = monthKey
    ws.Range("B4").value = fy

    Dim beginCash As Double: beginCash = GetMonthStatusValue(monthKey, "BeginningBalance")
    Dim endCash As Double: endCash = GetMonthStatusValue(monthKey, "EndingBalance")

    Dim inc As Double: inc = SumLedgerNet(monthKey, "Income", "(All)", "(All)", "(All)")
    Dim exp As Double: exp = SumLedgerNet(monthKey, "Expense", "(All)", "(All)", "(All)") + SumLedgerNet(monthKey, "Reimbursement", "(All)", "(All)", "(All)")
    Dim netChg As Double: netChg = inc - exp

    ws.Range("B7").value = beginCash
    ws.Range("B8").value = inc
    ws.Range("B9").value = exp
    ws.Range("B10").value = netChg
    ws.Range("B11").value = endCash

    ws.Range("E7").value = IIf(IsReconOk(monthKey), "YES", "NO")
    ws.Range("E8").value = IIf(IsMonthClosed(monthKey), "YES", "NO")

    Dim mc As Long, mr As Long, mAmt As Double, reconOk As Boolean, charityImbalance As Boolean, budgetOverrun As Boolean
    Call GateCheckMonth(monthKey, mc, mr, mAmt, reconOk, charityImbalance, budgetOverrun)
    ws.Range("E9").value = CStr(mr) & " / $" & Format$(mAmt, "0.00")
    ws.Range("E10").value = CStr(mc)

    Dim categories As Variant
    categories = GetCategoryList()

    Dim r As Long: r = 15
    Dim i As Long
    For i = LBound(categories) To UBound(categories)
        Dim cat As String: cat = CStr(categories(i))
        ws.Cells(r, 1).value = cat

        Dim b As Double: b = GetBudget(monthKey, cat)
        Dim a As Double: a = SumLedgerNet(monthKey, "(All)", "(All)", "(All)", cat)
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

    Dim charityRaisedM As Double: charityRaisedM = SumCharity(monthKey, "Raised")
    Dim charityPaidM As Double: charityPaidM = SumCharity(monthKey, "Paid")
    Dim charityHeldYTDVal As Double: charityHeldYTDVal = SumCharityYTD(monthKey, "Raised") - SumCharityYTD(monthKey, "Paid")

    ws.Range("H14").value = charityRaisedM
    ws.Range("H15").value = charityPaidM
    ws.Range("H16").value = charityHeldYTDVal
    ws.Range("H17").value = SumCharityYTD(monthKey, "Raised")
    ws.Range("H18").value = SumCharityYTD(monthKey, "Paid")

    WriteEventRollup ws, monthKey

    Dim ytdInc As Double: ytdInc = SumLedgerNetYTD(monthKey, "Income")
    Dim ytdExp As Double: ytdExp = SumLedgerNetYTD(monthKey, "Expense") + SumLedgerNetYTD(monthKey, "Reimbursement")
    ws.Range("B36").value = ytdInc
    ws.Range("B37").value = ytdExp
    ws.Range("B38").value = ytdInc - ytdExp

    Dim ytdBudgetVar As Double: ytdBudgetVar = BudgetVarYTDValue(monthKey)
    ws.Range("B39").value = ytdBudgetVar

    ws.Columns("A:H").AutoFit
    ws.Range("B7:B11,H14:H18,B36:B39").NumberFormat = "$#,##0.00"
    ws.Range("B15:D30").NumberFormat = "$#,##0.00"
    ws.Range("E15:E30").NumberFormat = "0.0%"

    Dim outFolder As String: outFolder = ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_BOARDPACKETS, ".\BoardPackets\"))
    EnsureFolderPath outFolder
    Dim pdfPath As String: pdfPath = outFolder & "TCPP_BoardPacket_" & monthKey & ".pdf"

    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    ArchiveBoardPacketSnapshot monthKey, ws, pdfPath

    ws.Visible = oldVis
    AuditLog "GenerateMonthlyPacket", "", pdfPath

CleanExit:
    Application.EnableEvents = prevEV
    Application.ScreenUpdating = prevSU
    Exit Sub
EH:
    HandleError "GenerateMonthlyPacket", Err, monthKey
    Resume CleanExit
End Sub

Private Sub ArchiveBoardPacketSnapshot(ByVal monthKey As String, ByVal wsReport As Worksheet, ByVal pdfPath As String)
    Dim archWs As Worksheet: Set archWs = GetSheet(SH_ARCHIVE)
    Dim archLo As ListObject: Set archLo = archWs.ListObjects("ARCH_BoardPackets")
    Dim lr As ListRow: Set lr = archLo.ListRows.Add

    lr.Range.Cells(1, archLo.ListColumns("MonthKey").Index).value = monthKey
    lr.Range.Cells(1, archLo.ListColumns("FiscalYear").Index).value = FiscalYearForMonthKey(monthKey)
    lr.Range.Cells(1, archLo.ListColumns("GeneratedAt").Index).value = Now
    lr.Range.Cells(1, archLo.ListColumns("GeneratedBy").Index).value = Application.UserName
    lr.Range.Cells(1, archLo.ListColumns("PdfPath").Index).value = pdfPath
    lr.Range.Cells(1, archLo.ListColumns("SnapshotRange").Index).value = "ARCH_BP_" & Replace(monthKey, "-", "_") & "!A1:H45"

    Dim snapName As String: snapName = "ARCH_BP_" & Replace(monthKey, "-", "_")
    Dim snapWs As Worksheet
    On Error Resume Next
    Set snapWs = ThisWorkbook.Worksheets(snapName)
    On Error GoTo 0
    If snapWs Is Nothing Then
        Set snapWs = ThisWorkbook.Worksheets.Add(After:=archWs)
        snapWs.name = snapName
    Else
        snapWs.Cells.ClearContents
    End If

    wsReport.Range("A1:H45").Copy
    snapWs.Range("A1").PasteSpecial xlPasteValues
    snapWs.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    lr.Range.Cells(1, archLo.ListColumns("SnapshotHash").Index).value = HashFile(pdfPath)
    snapWs.Visible = xlSheetVeryHidden
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
                If cat <> categoryFilter Then GoTo ContinueRow
                If LCase$(t) = "expense" Or LCase$(t) = "reimbursement" Then
                    SumLedgerNet = SumLedgerNet + Abs(CDbl(lo.DataBodyRange.Cells(i, netCol).value))
                End If
            Else
                If LCase$(t) = "income" Then
                    SumLedgerNet = SumLedgerNet + CDbl(lo.DataBodyRange.Cells(i, netCol).value)
                ElseIf LCase$(t) = "expense" Or LCase$(t) = "reimbursement" Then
                    SumLedgerNet = SumLedgerNet + Abs(CDbl(lo.DataBodyRange.Cells(i, netCol).value))
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
    y = CLng(left$(monthKey, 4))
    m = CLng(right$(monthKey, 2))
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

Public Function CharityHeldYTD(ByVal monthKey As String) As Double
    On Error GoTo EH
    CharityHeldYTD = SumCharityYTD(monthKey, "Raised") - SumCharityYTD(monthKey, "Paid")
    Exit Function
EH:
    HandleError "CharityHeldYTD", Err, monthKey
    CharityHeldYTD = 0#
End Function

Private Function CharityPaidExceedsRaisedYTD(ByVal monthKey As String) As Boolean
    CharityPaidExceedsRaisedYTD = (SumCharityYTD(monthKey, "Paid") > SumCharityYTD(monthKey, "Raised"))
End Function

Private Sub WriteEventRollup(ByVal ws As Worksheet, ByVal monthKey As String)
    ws.Range("A26:D34").ClearContents

    Dim evLo As ListObject: Set evLo = GetTable(SH_LOOKUPS, T_EVENTS_LIST)
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
    On Error GoTo EH
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
    Exit Function
EH:
    HandleError "GetBudget", Err, monthKey & "|" & category
    GetBudget = 0#
End Function

Public Sub SetBudget(ByVal monthKey As String, ByVal category As String, ByVal amount As Double)
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_BUDGET, T_BUDGET)
    Dim rowIdx As Long
    rowIdx = 0

    If Not lo.DataBodyRange Is Nothing Then
        Dim mkCol As Long: mkCol = lo.ListColumns("MonthKey").Index
        Dim catCol As Long: catCol = lo.ListColumns("Category").Index
        Dim i As Long
        For i = 1 To lo.ListRows.count
            If CStr(lo.DataBodyRange.Cells(i, mkCol).value) = monthKey And _
               CStr(lo.DataBodyRange.Cells(i, catCol).value) = category Then
                rowIdx = i
                Exit For
            End If
        Next i
    End If

    If rowIdx = 0 Then
        Dim lr As ListRow: Set lr = lo.ListRows.Add
        rowIdx = lr.Index
        lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("MonthKey").Index).value = monthKey
        lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("FiscalYear").Index).value = FiscalYearForMonthKey(monthKey)
        lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("Category").Index).value = category
    End If

    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("BudgetAmount").Index).value = amount
    Exit Sub
EH:
    HandleError "SetBudget", Err, monthKey & "|" & category
End Sub

Public Function BudgetVarYTDValue(ByVal monthKey As String) As Double
    On Error GoTo EH
    Dim startKey As String: startKey = FiscalYearStartMonthKey(monthKey)
    Dim cur As String: cur = startKey
    Dim total As Double: total = 0#

    Do While MonthKeyLessOrEqual(cur, monthKey)
        total = total + BudgetVarMonthValue(cur)
        cur = MonthKeyAdd(cur, 1)
    Loop

    BudgetVarYTDValue = total
    Exit Function
EH:
    HandleError "BudgetVarYTDValue", Err, monthKey
    BudgetVarYTDValue = 0#
End Function

Private Function BudgetVarMonthValue(ByVal monthKey As String) As Double
    Dim categories As Variant
    categories = GetCategoryList()

    Dim i As Long
    For i = LBound(categories) To UBound(categories)
        Dim cat As String: cat = CStr(categories(i))
        BudgetVarMonthValue = BudgetVarMonthValue + (GetBudget(monthKey, cat) - SumLedgerNet(monthKey, "(All)", "(All)", "(All)", cat))
    Next i
End Function

Private Function HasBudgetOverrun(ByVal monthKey As String) As Boolean
    Dim categories As Variant
    categories = GetCategoryList()

    Dim i As Long
    For i = LBound(categories) To UBound(categories)
        Dim cat As String: cat = CStr(categories(i))
        Dim b As Double: b = GetBudget(monthKey, cat)
        Dim a As Double: a = SumLedgerNet(monthKey, "(All)", "(All)", "(All)", cat)
        If b > 0 And a > b Then
            HasBudgetOverrun = True
            Exit Function
        End If
    Next i
    HasBudgetOverrun = False
End Function

Private Function GetCategoryList() As Variant
    Dim lo As ListObject: Set lo = GetTable(SH_LOOKUPS, T_COA)
    Dim items() As String
    Dim count As Long

    If lo.DataBodyRange Is Nothing Then
        GetCategoryList = Array("Administrative", "Programs", "Fundraising", "Marketing", "Travel", "Services", "Misc")
        Exit Function
    End If

    Dim cell As Range
    For Each cell In lo.ListColumns(1).DataBodyRange.Cells
        If Len(Trim$(CStr(cell.value))) > 0 Then
            ReDim Preserve items(count)
            items(count) = CStr(cell.value)
            count = count + 1
        End If
    Next cell

    If count = 0 Then
        GetCategoryList = Array("Administrative", "Programs", "Fundraising", "Marketing", "Travel", "Services", "Misc")
    Else
        GetCategoryList = items
    End If
End Function

Private Sub GetReceiptExceptions(ByVal monthKey As String, ByRef missingReceipts As Long, ByRef missingReceiptAmt As Double)
    missingReceipts = 0
    missingReceiptAmt = 0#

    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim mkCol As Long: mkCol = lo.ListColumns("MonthKey").Index
    Dim rrCol As Long: rrCol = lo.ListColumns("ReceiptRequired").Index
    Dim rsCol As Long: rsCol = lo.ListColumns("ReceiptStatus").Index
    Dim netCol As Long: netCol = lo.ListColumns("Net").Index

    Dim i As Long
    For i = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(i, mkCol).value) = monthKey Then
            If CBool(lo.DataBodyRange.Cells(i, rrCol).value) Then
                Dim rs As String: rs = CStr(lo.DataBodyRange.Cells(i, rsCol).value)
                If rs <> "Recorded" And rs <> "Waived" Then
                    missingReceipts = missingReceipts + 1
                    missingReceiptAmt = missingReceiptAmt + Abs(CDbl(lo.DataBodyRange.Cells(i, netCol).value))
                End If
            End If
        End If
    Next i
End Sub

Private Function GetUncategorizedCount(ByVal monthKey As String) As Long
    GetUncategorizedCount = 0
    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim mkCol As Long: mkCol = lo.ListColumns("MonthKey").Index
    Dim catCol As Long: catCol = lo.ListColumns("Category").Index

    Dim i As Long
    For i = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(i, mkCol).value) = monthKey Then
            If Len(Trim$(CStr(lo.DataBodyRange.Cells(i, catCol).value))) = 0 Then
                GetUncategorizedCount = GetUncategorizedCount + 1
            End If
        End If
    Next i
End Function

'========================
' Exceptions / dashboard support
'========================

Public Sub GetExceptionCounts(ByVal monthKey As String, ByRef uncategorized As Long, ByRef missingReceipts As Long, ByRef missingReceiptAmt As Double)
    On Error GoTo EH
    Dim reconOk As Boolean, charityImbalance As Boolean, budgetOverrun As Boolean
    Dim msg As String
    msg = GateCheckMonth(monthKey, uncategorized, missingReceipts, missingReceiptAmt, reconOk, charityImbalance, budgetOverrun)
    Exit Sub
EH:
    HandleError "GetExceptionCounts", Err, monthKey
End Sub

Public Sub GetDashboardMetrics(ByVal monthKey As String, ByVal eventFilter As String, ByVal charityFilter As String, _
                               ByRef totalIncome As Double, ByRef totalExpense As Double, ByRef netChange As Double, _
                               ByRef missingReceiptCount As Long, ByRef missingReceiptAmt As Double, _
                               ByRef uncategorizedCount As Long, ByRef charityRaised As Double, ByRef charityPaid As Double, _
                               ByRef charityHeld As Double, ByRef budgetVarMonth As Double, ByRef budgetVarYTD As Double)
    On Error GoTo EH
    totalIncome = SumLedgerNet(monthKey, "Income", eventFilter, charityFilter, "(All)")
    totalExpense = SumLedgerNet(monthKey, "Expense", eventFilter, charityFilter, "(All)") + _
                   SumLedgerNet(monthKey, "Reimbursement", eventFilter, charityFilter, "(All)")
    netChange = totalIncome - totalExpense

    GetReceiptExceptions monthKey, missingReceiptCount, missingReceiptAmt
    uncategorizedCount = GetUncategorizedCount(monthKey)

    charityRaised = SumCharity(monthKey, "Raised")
    charityPaid = SumCharity(monthKey, "Paid")
    charityHeld = CharityHeldYTD(monthKey)

    budgetVarMonth = BudgetVarMonthValue(monthKey)
    budgetVarYTD = BudgetVarYTDValue(monthKey)
    Exit Sub
EH:
    HandleError "GetDashboardMetrics", Err, monthKey
End Sub

'========================
' Minutes + Agenda (Word automation)
'========================

Public Function CreateMeeting(ByVal meetingDate As Date, ByVal scribe As String, ByVal location As String) As String
    On Error GoTo EH
    Dim meetingId As String
    meetingId = "MTG-" & Format$(meetingDate, "yyyymmdd") & "-" & Format$(Int((9999 * Rnd) + 1), "0000")

    Dim docPath As String
    docPath = ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_MINUTES_DOCX, ".\Minutes\DOCX\")) & "TCPP_Minutes_" & Format$(meetingDate, "yyyy-mm-dd") & ".docx"

    Dim templatePath As String
    templatePath = GetTemplatePath("TCPP Board Meeting Minutes Template.docx")
    If Len(Dir(templatePath)) = 0 Then
        Err.Raise vbObjectError + 704, "CreateMeeting", "Minutes template not found: " & templatePath
    End If
    FileCopy templatePath, docPath

    Dim lo As ListObject: Set lo = GetTable(SH_MEETINGS, T_MEETINGS)
    Dim lr As ListRow: Set lr = lo.ListRows.Add

    lr.Range.Cells(1, lo.ListColumns("MeetingID").Index).value = meetingId
    lr.Range.Cells(1, lo.ListColumns("MeetingDate").Index).value = meetingDate
    lr.Range.Cells(1, lo.ListColumns("StartTime").Index).value = ""
    lr.Range.Cells(1, lo.ListColumns("EndTime").Index).value = ""
    lr.Range.Cells(1, lo.ListColumns("Scribe").Index).value = scribe
    lr.Range.Cells(1, lo.ListColumns("Location").Index).value = location
    lr.Range.Cells(1, lo.ListColumns("MeetingType").Index).value = "Regular"
    lr.Range.Cells(1, lo.ListColumns("OpenSessionFlag").Index).value = True
    lr.Range.Cells(1, lo.ListColumns("RemoteAllowedFlag").Index).value = False
    lr.Range.Cells(1, lo.ListColumns("MinutesDocPath").Index).value = docPath
    lr.Range.Cells(1, lo.ListColumns("MinutesPdfPath").Index).value = ""
    lr.Range.Cells(1, lo.ListColumns("CreatedAt").Index).value = Now

    OpenWordDocument docPath
    AuditLog "CreateMeeting", meetingId, docPath
    CreateMeeting = meetingId
    Exit Function
EH:
    HandleError "CreateMeeting", Err, Format$(meetingDate, "yyyy-mm-dd")
    CreateMeeting = ""
End Function

Public Sub ExportMeetingPdf(ByVal meetingId As String)
    Dim prevSU As Boolean, prevEV As Boolean
    prevSU = Application.ScreenUpdating
    prevEV = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_MEETINGS, T_MEETINGS)
    Dim rowIdx As Long: rowIdx = FindMeetingRowIndex(meetingId)
    If rowIdx = 0 Then Err.Raise vbObjectError + 702, "ExportMeetingPdf", "MeetingID not found"

    Dim meetingDate As Date
    meetingDate = CDate(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("MeetingDate").Index).value)

    Dim docPath As String
    docPath = CStr(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("MinutesDocPath").Index).value)
    Dim pdfPath As String
    pdfPath = ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_MINUTES_PDF, ".\Minutes\PDF\")) & "TCPP_Minutes_" & Format$(meetingDate, "yyyy-mm-dd") & ".pdf"

    ExportWordPdf docPath, pdfPath
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("MinutesPdfPath").Index).value = pdfPath

    AuditLog "ExportMeetingPdf", meetingId, pdfPath
    GoTo CleanExit
EH:
    HandleError "ExportMeetingPdf", Err, meetingId
CleanExit:
    Application.EnableEvents = prevEV
    Application.ScreenUpdating = prevSU
End Sub

Public Function CreateAgenda(ByVal agendaDate As Date) As String
    On Error GoTo EH
    Dim agendaId As String
    agendaId = "AGD-" & Format$(agendaDate, "yyyymmdd") & "-" & Format$(Int((9999 * Rnd) + 1), "0000")

    Dim docPath As String
    docPath = ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_AGENDA_DOCX, ".\Agenda\DOCX\")) & "TCPP_Agenda_" & Format$(agendaDate, "yyyy-mm-dd") & ".docx"

    Dim templatePath As String
    templatePath = GetTemplatePath("Template Meeting Agenda.docx")
    If Len(Dir(templatePath)) = 0 Then
        Err.Raise vbObjectError + 705, "CreateAgenda", "Agenda template not found: " & templatePath
    End If
    FileCopy templatePath, docPath

    Dim lo As ListObject: Set lo = GetTable(SH_AGENDA, T_AGENDA)
    Dim lr As ListRow: Set lr = lo.ListRows.Add

    lr.Range.Cells(1, lo.ListColumns("AgendaID").Index).value = agendaId
    lr.Range.Cells(1, lo.ListColumns("AgendaDate").Index).value = agendaDate
    lr.Range.Cells(1, lo.ListColumns("DocPath").Index).value = docPath
    lr.Range.Cells(1, lo.ListColumns("PdfPath").Index).value = ""
    lr.Range.Cells(1, lo.ListColumns("CreatedAt").Index).value = Now

    OpenWordDocument docPath
    AuditLog "CreateAgenda", agendaId, docPath
    CreateAgenda = agendaId
    Exit Function
EH:
    HandleError "CreateAgenda", Err, Format$(agendaDate, "yyyy-mm-dd")
    CreateAgenda = ""
End Function

Public Sub ExportAgendaPdf(ByVal agendaId As String)
    Dim prevSU As Boolean, prevEV As Boolean
    prevSU = Application.ScreenUpdating
    prevEV = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_AGENDA, T_AGENDA)
    Dim rowIdx As Long: rowIdx = FindAgendaRowIndex(agendaId)
    If rowIdx = 0 Then Err.Raise vbObjectError + 703, "ExportAgendaPdf", "AgendaID not found"

    Dim agendaDate As Date
    agendaDate = CDate(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("AgendaDate").Index).value)

    Dim docPath As String
    docPath = CStr(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("DocPath").Index).value)
    Dim pdfPath As String
    pdfPath = ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_AGENDA_PDF, ".\Agenda\PDF\")) & "TCPP_Agenda_" & Format$(agendaDate, "yyyy-mm-dd") & ".pdf"

    ExportWordPdf docPath, pdfPath
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("PdfPath").Index).value = pdfPath

    AuditLog "ExportAgendaPdf", agendaId, pdfPath
    GoTo CleanExit
EH:
    HandleError "ExportAgendaPdf", Err, agendaId
CleanExit:
    Application.EnableEvents = prevEV
    Application.ScreenUpdating = prevSU
End Sub

Private Function FindMeetingRowIndex(ByVal meetingId As String) As Long
    Dim lo As ListObject: Set lo = GetTable(SH_MEETINGS, T_MEETINGS)
    FindMeetingRowIndex = 0
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Range
    For Each r In lo.ListColumns("MeetingID").DataBodyRange.Cells
        If CStr(r.value) = meetingId Then
            FindMeetingRowIndex = r.Row - lo.DataBodyRange.Row + 1
            Exit Function
        End If
    Next r
End Function

Private Function FindAgendaRowIndex(ByVal agendaId As String) As Long
    Dim lo As ListObject: Set lo = GetTable(SH_AGENDA, T_AGENDA)
    FindAgendaRowIndex = 0
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Range
    For Each r In lo.ListColumns("AgendaID").DataBodyRange.Cells
        If CStr(r.value) = agendaId Then
            FindAgendaRowIndex = r.Row - lo.DataBodyRange.Row + 1
            Exit Function
        End If
    Next r
End Function

Private Function GetTemplatePath(ByVal templateName As String) As String
    Dim basePath As String: basePath = ThisWorkbook.path
    If Len(basePath) = 0 Then basePath = CurDir$
    GetTemplatePath = basePath & "\" & templateName
End Function

Private Sub OpenWordDocument(ByVal docPath As String)
    Dim wdApp As Object
    If Len(Dir(docPath)) = 0 Then Err.Raise vbObjectError + 706, "OpenWordDocument", "Doc not found: " & docPath
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0

    wdApp.Visible = True
    wdApp.Documents.Open docPath
End Sub

Private Sub ExportWordPdf(ByVal docPath As String, ByVal pdfPath As String)
    Dim wdApp As Object, doc As Object
    On Error GoTo EH
    If Len(Dir(docPath)) = 0 Then Err.Raise vbObjectError + 707, "ExportWordPdf", "Doc not found: " & docPath
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set doc = wdApp.Documents.Open(docPath)
    doc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=17
    doc.Close SaveChanges:=False
    wdApp.Quit
    Exit Sub
EH:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    On Error GoTo 0
    Err.Raise Err.Number, "ExportWordPdf", Err.Description
End Sub

'========================
' Members
'========================

Public Sub UpsertMember(ByVal memberName As String, ByVal memberEmail As String, ByVal membershipType As String, _
                        ByVal duesPaid As Boolean, ByVal duesPaidDate As Variant, ByVal duesAmount As Double, _
                        ByVal joinedDate As Variant, ByVal notes As String, Optional ByVal externalSource As String = "Manual", _
                        Optional ByVal externalMemberId As String = "")
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_MEMBERS, T_MEMBERS)
    Dim rowIdx As Long: rowIdx = FindMemberRowIndex(memberEmail, memberName)

    If rowIdx = 0 Then
        Dim lr As ListRow: Set lr = lo.ListRows.Add
        rowIdx = lr.Index
        If Len(memberEmail) > 0 Then
            lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("MemberID").Index).value = "MBR-" & LCase$(memberEmail)
        Else
            lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("MemberID").Index).value = "MBR-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(Int((9999 * Rnd) + 1), "0000")
        End If
    End If

    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("MemberName").Index).value = memberName
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("MemberEmail").Index).value = memberEmail
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("MembershipType").Index).value = membershipType
    If IsDate(joinedDate) Then lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("JoinedDate").Index).value = CDate(joinedDate)
    If ColumnExists(lo, "Status") And Len(CStr(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("Status").Index).value)) = 0 Then
        lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("Status").Index).value = IIf(duesPaid, "Active", "ApprovedPendingDues")
    End If
    If ColumnExists(lo, "RenewalIntervalMonths") Then
        lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("RenewalIntervalMonths").Index).value = CLng(GetConfigValue(CFG_RENEWAL_INTERVAL, "12"))
    End If
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("DuesPaidFlag").Index).value = IIf(duesPaid, "Y", "N")
    If IsDate(duesPaidDate) Then lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("DuesPaidDate").Index).value = CDate(duesPaidDate)
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("DuesAmount").Index).value = duesAmount
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("RenewalDate").Index).value = CalculateRenewalDate(duesPaidDate)
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("Notes").Index).value = notes
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ExternalSource").Index).value = externalSource
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ExternalMemberID").Index).value = externalMemberId
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("LastUpdatedAt").Index).value = Now
    Exit Sub
EH:
    HandleError "UpsertMember", Err, memberEmail
End Sub

Private Function FindMemberIdByEmail(ByVal memberEmail As String) As String
    FindMemberIdByEmail = ""
    If Len(memberEmail) = 0 Then Exit Function
    Dim lo As ListObject: Set lo = GetTable(SH_MEMBERS, T_MEMBERS)
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Range
    For Each r In lo.ListColumns("MemberEmail").DataBodyRange.Cells
        If LCase$(CStr(r.value)) = LCase$(memberEmail) Then
            FindMemberIdByEmail = CStr(r.Offset(0, lo.ListColumns("MemberID").Index - lo.ListColumns("MemberEmail").Index).value)
            Exit Function
        End If
    Next r
End Function

Public Function CalculateRenewalDate(ByVal duesPaidDate As Variant) As Variant
    On Error GoTo EH
    If Not IsDate(duesPaidDate) Then
        CalculateRenewalDate = ""
        Exit Function
    End If

    Dim months As Long
    months = CLng(GetConfigValue(CFG_RENEWAL_INTERVAL, "12"))
    CalculateRenewalDate = DateAdd("m", months, CDate(duesPaidDate))
    Exit Function
EH:
    HandleError "CalculateRenewalDate", Err, ""
    CalculateRenewalDate = ""
End Function

Private Function FindMemberRowIndex(ByVal email As String, ByVal name As String) As Long
    Dim lo As ListObject: Set lo = GetTable(SH_MEMBERS, T_MEMBERS)
    FindMemberRowIndex = 0
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Range
    If Len(email) > 0 Then
        For Each r In lo.ListColumns("MemberEmail").DataBodyRange.Cells
            If LCase$(CStr(r.value)) = LCase$(email) Then
                FindMemberRowIndex = r.Row - lo.DataBodyRange.Row + 1
                Exit Function
            End If
        Next r
    End If

    For Each r In lo.ListColumns("MemberName").DataBodyRange.Cells
        If LCase$(CStr(r.value)) = LCase$(name) Then
            FindMemberRowIndex = r.Row - lo.DataBodyRange.Row + 1
            Exit Function
        End If
    Next r
End Function

'========================
' Imports (scaffold)
'========================

Public Sub ImportCsvRaw(ByVal sourceName As String, ByVal filePath As String)
    Dim prevSU As Boolean, prevEV As Boolean
    prevSU = Application.ScreenUpdating
    prevEV = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo EH
    If Len(Dir(filePath)) = 0 Then Err.Raise vbObjectError + 708, "ImportCsvRaw", "File not found: " & filePath

    Dim batchId As String
    batchId = "IMP-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(Int((9999 * Rnd) + 1), "0000")

    Dim fileHash As String
    fileHash = HashFile(filePath)

    Dim rawTable As ListObject
    If LCase$(sourceName) = "zeffy" Then
        Set rawTable = GetTable(SH_IMPORTS, T_ZEFFY_RAW)
    Else
        Set rawTable = GetTable(SH_IMPORTS, T_BLAZE_RAW)
    End If

    Dim rowCount As Long
    Dim f As Integer: f = FreeFile
    Open filePath For Input As #f

    Dim headerLine As String
    Dim headers() As String
    Dim hasHeader As Boolean

    If Not EOF(f) Then
        Line Input #f, headerLine
        headers = ParseCsvLine(headerLine)
        hasHeader = True
    End If

    Dim mapDict As Object
    Set mapDict = GetImportMapping(sourceName)

    Dim line As String
    Dim rowNum As Long
    Do While Not EOF(f)
        Line Input #f, line
        If Len(Trim$(line)) = 0 Then GoTo ContinueRow

        rowNum = rowNum + 1
        Dim rowHash As String: rowHash = HashRow(line)
        Dim data() As String
        If hasHeader Then data = ParseCsvLine(line)

        Dim extId As String
        extId = ""
        If hasHeader Then extId = ExtractMappedValue(mapDict, headers, data, "ExternalTxnID")

        If Len(extId) > 0 Then
            If LedgerHasExternalTxn(sourceName, extId) Then GoTo ContinueRow
        ElseIf RawRowExists(rawTable, rowHash) Then
            GoTo ContinueRow
        End If

        Dim lr As ListRow: Set lr = rawTable.ListRows.Add
        SetIfExists rawTable, lr, "ImportBatchID", batchId
        SetIfExists rawTable, lr, "RowNum", rowNum
        SetIfExists rawTable, lr, "RawLine", line
        SetIfExists rawTable, lr, "ExternalTxnID", extId
        SetIfExists rawTable, lr, "RowHash", rowHash
        SetIfExists rawTable, lr, "CreatedAt", Now
        rowCount = rowCount + 1

        If mapDict.count > 0 Then
            TryMapRowToLedger sourceName, batchId, headers, data, mapDict
        End If
ContinueRow:
    Loop
    Close #f

    Dim log As ListObject: Set log = GetTable(SH_IMPORTS, T_IMPORTLOG)
    Dim lrLog As ListRow: Set lrLog = log.ListRows.Add
    lrLog.Range.Cells(1, log.ListColumns("ImportBatchID").Index).value = batchId
    lrLog.Range.Cells(1, log.ListColumns("Source").Index).value = sourceName
    lrLog.Range.Cells(1, log.ListColumns("ImportedAt").Index).value = Now
    lrLog.Range.Cells(1, log.ListColumns("FileName").Index).value = filePath
    lrLog.Range.Cells(1, log.ListColumns("FileHash").Index).value = fileHash
    lrLog.Range.Cells(1, log.ListColumns("RowCount").Index).value = rowCount
    lrLog.Range.Cells(1, log.ListColumns("Notes").Index).value = "Raw staging import"
    lrLog.Range.Cells(1, log.ListColumns("Status").Index).value = "OK"

    Dim archivePath As String
    archivePath = ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_IMPORTS_ARCHIVE, ".\\Imports\\Archive\\")) & _
        sourceName & "_" & batchId & "_" & Dir(filePath)
    FileCopy filePath, archivePath

    AuditLog "ImportCsvRaw", batchId, sourceName & " rows=" & CStr(rowCount)
    GoTo CleanExit
EH:
    HandleError "ImportCsvRaw", Err, filePath
CleanExit:
    Application.EnableEvents = prevEV
    Application.ScreenUpdating = prevSU
End Sub

Public Sub ExportDuesReport()
    Dim prevSU As Boolean, prevEV As Boolean
    prevSU = Application.ScreenUpdating
    prevEV = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo EH

    Dim outFolder As String
    outFolder = ResolveWorkbookRelativePath(".\\Reports\\")
    EnsureFolderPath outFolder

    Dim outPath As String
    outPath = outFolder & "TCPP_DuesStatus_" & Format$(Date, "yyyymmdd") & ".csv"

    Dim lo As ListObject: Set lo = GetTable(SH_MEMBERS, T_MEMBERS)
    Dim f As Integer: f = FreeFile
    Open outPath For Output As #f

    Print #f, "MemberName,MemberEmail,MembershipType,DuesPaidFlag,DuesPaidDate,DuesAmount,RenewalDate,Notes"
    If Not lo.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = 1 To lo.ListRows.count
            Print #f, CsvCell(lo.DataBodyRange.Cells(i, lo.ListColumns("MemberName").Index).value) & "," & _
                      CsvCell(lo.DataBodyRange.Cells(i, lo.ListColumns("MemberEmail").Index).value) & "," & _
                      CsvCell(lo.DataBodyRange.Cells(i, lo.ListColumns("MembershipType").Index).value) & "," & _
                      CsvCell(lo.DataBodyRange.Cells(i, lo.ListColumns("DuesPaidFlag").Index).value) & "," & _
                      CsvCell(lo.DataBodyRange.Cells(i, lo.ListColumns("DuesPaidDate").Index).value) & "," & _
                      CsvCell(lo.DataBodyRange.Cells(i, lo.ListColumns("DuesAmount").Index).value) & "," & _
                      CsvCell(lo.DataBodyRange.Cells(i, lo.ListColumns("RenewalDate").Index).value) & "," & _
                      CsvCell(lo.DataBodyRange.Cells(i, lo.ListColumns("Notes").Index).value)
        Next i
    End If

    Close #f
    AuditLog "ExportDuesReport", "", outPath
    MsgBox "Dues report exported: " & outPath, vbInformation, "TCPP"

CleanExit:
    Application.EnableEvents = prevEV
    Application.ScreenUpdating = prevSU
    Exit Sub
EH:
    HandleError "ExportDuesReport", Err, ""
    Resume CleanExit
End Sub

Private Function CsvCell(ByVal v As Variant) As String
    Dim s As String: s = CStr(v)
    s = Replace(s, """", """""")
    CsvCell = """" & s & """"
End Function

Private Function GetImportMapping(ByVal sourceName As String) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Dim mapTable As ListObject
    If LCase$(sourceName) = "zeffy" Then
        Set mapTable = GetTable(SH_IMPORTS, T_IMPORTMAP_ZEFFY)
    Else
        Set mapTable = GetTable(SH_IMPORTS, T_IMPORTMAP_BLAZE)
    End If

    If mapTable.DataBodyRange Is Nothing Then
        Set GetImportMapping = dict
        Exit Function
    End If

    Dim i As Long
    For i = 1 To mapTable.ListRows.count
        Dim activeFlag As String: activeFlag = "TRUE"
        If ColumnExists(mapTable, "ActiveFlag") Then
            activeFlag = CStr(mapTable.DataBodyRange.Cells(i, mapTable.ListColumns("ActiveFlag").Index).value)
        End If
        If LCase$(activeFlag) = "false" Then GoTo ContinueMap

        Dim src As String: src = CStr(mapTable.DataBodyRange.Cells(i, mapTable.ListColumns("SourceColumn").Index).value)
        Dim tgtTable As String: tgtTable = ""
        If ColumnExists(mapTable, "TargetTable") Then
            tgtTable = CStr(mapTable.DataBodyRange.Cells(i, mapTable.ListColumns("TargetTable").Index).value)
        End If
        Dim tgtCol As String: tgtCol = CStr(mapTable.DataBodyRange.Cells(i, mapTable.ListColumns("TargetColumn").Index).value)
        If Len(tgtTable) = 0 Then tgtTable = T_LEDGER
        If Len(src) > 0 And Len(tgtCol) > 0 Then
            dict(LCase$(tgtTable & "|" & tgtCol)) = src
        End If
ContinueMap:
    Next i

    Set GetImportMapping = dict
End Function

Private Function ExtractMappedValue(ByVal mapDict As Object, ByVal headers() As String, ByVal data() As String, ByVal targetCol As String) As String
    Dim srcCol As String
    ExtractMappedValue = ""
    Dim key As String: key = LCase$(T_LEDGER & "|" & targetCol)
    If mapDict.Exists(key) Then
        srcCol = mapDict(key)
        Dim idx As Long: idx = HeaderIndex(headers, srcCol)
        If idx >= 0 And idx <= UBound(data) Then
            ExtractMappedValue = data(idx)
        End If
    End If
End Function

Private Sub TryMapRowToLedger(ByVal sourceName As String, ByVal batchId As String, ByVal headers() As String, ByVal data() As String, ByVal mapDict As Object)
    On Error GoTo EH

    Dim txnDate As Date
    Dim txnType As String
    Dim txnSubtype As String
    Dim category As String
    Dim eventName As String
    Dim charityName As String
    Dim gross As Double
    Dim fees As Double
    Dim paymentMethod As String
    Dim sourceType As String
    Dim sourceNameVal As String
    Dim memberName As String
    Dim memberEmail As String
    Dim memo As String
    Dim extId As String

    txnDate = CDate(ExtractMappedValue(mapDict, headers, data, "Date"))
    txnType = ExtractMappedValue(mapDict, headers, data, "TxnType")
    txnSubtype = ExtractMappedValue(mapDict, headers, data, "TxnSubtype")
    category = ExtractMappedValue(mapDict, headers, data, "Category")
    eventName = ExtractMappedValue(mapDict, headers, data, "Event")
    charityName = ExtractMappedValue(mapDict, headers, data, "Charity")
    gross = CDbl(Val(ExtractMappedValue(mapDict, headers, data, "Gross")))
    fees = CDbl(Val(ExtractMappedValue(mapDict, headers, data, "Fees")))
    paymentMethod = ExtractMappedValue(mapDict, headers, data, "PaymentMethod")
    sourceType = ExtractMappedValue(mapDict, headers, data, "SourceType")
    sourceNameVal = ExtractMappedValue(mapDict, headers, data, "SourceName")
    memberName = ExtractMappedValue(mapDict, headers, data, "MemberName")
    memberEmail = ExtractMappedValue(mapDict, headers, data, "MemberEmail")
    memo = ExtractMappedValue(mapDict, headers, data, "Memo")
    extId = ExtractMappedValue(mapDict, headers, data, "ExternalTxnID")

    If Len(txnType) = 0 Or txnDate = 0 Then Exit Sub

    AddLedgerEntry txnDate, txnType, txnSubtype, category, eventName, charityName, gross, fees, paymentMethod, _
        sourceType, sourceNameVal, memberName, memberEmail, memo, False, sourceName, extId, batchId, False

    If LCase$(sourceName) = "zeffy" And Len(memberEmail) > 0 Then
        UpsertMember memberName, memberEmail, ExtractMappedValue(mapDict, headers, data, "MembershipType"), _
            True, txnDate, gross, txnDate, "Imported from Zeffy", "Zeffy", ExtractMappedValue(mapDict, headers, data, "ExternalMemberID")
    End If

    Exit Sub
EH:
    LogError "TryMapRowToLedger", Err, sourceName & " batch=" & batchId
End Sub

Private Function LedgerHasExternalTxn(ByVal sourceName As String, ByVal extId As String) As Boolean
    LedgerHasExternalTxn = False
    Dim lo As ListObject: Set lo = GetTable(SH_LEDGER, T_LEDGER)
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim sourceCol As Long: sourceCol = lo.ListColumns("ExternalSource").Index
    Dim extCol As Long: extCol = lo.ListColumns("ExternalTxnID").Index
    Dim i As Long
    For i = 1 To lo.ListRows.count
        If LCase$(CStr(lo.DataBodyRange.Cells(i, sourceCol).value)) = LCase$(sourceName) And _
           CStr(lo.DataBodyRange.Cells(i, extCol).value) = extId Then
            LedgerHasExternalTxn = True
            Exit Function
        End If
    Next i
End Function

Private Function HeaderIndex(ByVal headers() As String, ByVal name As String) As Long
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If LCase$(Trim$(headers(i))) = LCase$(Trim$(name)) Then
            HeaderIndex = i
            Exit Function
        End If
    Next i
    HeaderIndex = -1
End Function

Private Function ParseCsvLine(ByVal line As String) As String()
    Dim result() As String
    Dim idx As Long: idx = 0
    Dim i As Long
    Dim inQuotes As Boolean
    Dim token As String
    For i = 1 To Len(line)
        Dim ch As String: ch = Mid$(line, i, 1)
        If ch = """" Then
            If inQuotes And i < Len(line) And Mid$(line, i + 1, 1) = """" Then
                token = token & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = "," And Not inQuotes Then
            ReDim Preserve result(idx)
            result(idx) = token
            idx = idx + 1
            token = ""
        Else
            token = token & ch
        End If
    Next i
    ReDim Preserve result(idx)
    result(idx) = token
    ParseCsvLine = result
End Function

Private Function RawRowExists(ByVal rawTable As ListObject, ByVal rowHash As String) As Boolean
    RawRowExists = False
    If rawTable.DataBodyRange Is Nothing Then Exit Function

    Dim r As Range
    For Each r In rawTable.ListColumns("RowHash").DataBodyRange.Cells
        If CStr(r.value) = rowHash Then
            RawRowExists = True
            Exit Function
        End If
    Next r
End Function

Private Function Crc32(ByVal text As String) As Long
    Dim i As Long, j As Long
    Dim crc As Long
    crc = &HFFFFFFFF

    For i = 1 To Len(text)
        crc = crc Xor Asc(Mid$(text, i, 1))
        For j = 1 To 8
            If (crc And 1) Then
                crc = (crc \\ 2) Xor &HEDB88320
            Else
                crc = crc \\ 2
            End If
        Next j
    Next i
    Crc32 = Not crc
End Function

'========================
' Compliance
'========================

Public Function AddComplianceEvent(ByVal eventName As String, ByVal eventDate As Date, ByVal location As String, _
                                   ByVal eventType As String, ByVal ageRestriction As String, _
                                   ByVal charityPartner As String, ByVal purpose As String, _
                                   ByVal safetyLead As String, ByVal financeLead As String, _
                                   ByVal paymentMethods As String, Optional ByVal approvalStatus As String = "Draft") As String
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_COMPLIANCE, T_COMPLIANCE_EVENTS)
    Dim lr As ListRow: Set lr = lo.ListRows.Add

    Dim eventId As String
    eventId = "EVT-" & Format$(eventDate, "yyyymmdd") & "-" & Format$(Int((9999 * Rnd) + 1), "0000")

    lr.Range.Cells(1, lo.ListColumns("EventID").Index).value = eventId
    lr.Range.Cells(1, lo.ListColumns("EventName").Index).value = eventName
    lr.Range.Cells(1, lo.ListColumns("EventDate").Index).value = eventDate
    lr.Range.Cells(1, lo.ListColumns("Location").Index).value = location
    lr.Range.Cells(1, lo.ListColumns("EventType").Index).value = eventType
    lr.Range.Cells(1, lo.ListColumns("AgeRestriction").Index).value = ageRestriction
    lr.Range.Cells(1, lo.ListColumns("CharityPartner").Index).value = charityPartner
    lr.Range.Cells(1, lo.ListColumns("Purpose").Index).value = purpose
    lr.Range.Cells(1, lo.ListColumns("SafetyLead").Index).value = safetyLead
    lr.Range.Cells(1, lo.ListColumns("FinanceLead").Index).value = financeLead
    lr.Range.Cells(1, lo.ListColumns("PaymentMethodsOnSite").Index).value = paymentMethods
    lr.Range.Cells(1, lo.ListColumns("ApprovalStatus").Index).value = approvalStatus
    lr.Range.Cells(1, lo.ListColumns("CreatedAt").Index).value = Now
    lr.Range.Cells(1, lo.ListColumns("UpdatedAt").Index).value = Now

    AuditLog "ComplianceEvent", eventId, eventName
    AddComplianceEvent = eventId
    Exit Function
EH:
    HandleError "AddComplianceEvent", Err, eventName
    AddComplianceEvent = ""
End Function

Public Function AddRefundRequest(ByVal requestDate As Date, ByVal requesterName As String, ByVal requesterEmail As String, _
                                 ByVal amountRequested As Double, ByVal reason As String, _
                                 Optional ByVal txnId As String = "", Optional ByVal eventId As String = "") As String
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_COMPLIANCE, T_COMPLIANCE_REFUNDS)
    Dim lr As ListRow: Set lr = lo.ListRows.Add

    Dim refundId As String
    refundId = "RFD-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(Int((9999 * Rnd) + 1), "0000")

    lr.Range.Cells(1, lo.ListColumns("RefundID").Index).value = refundId
    lr.Range.Cells(1, lo.ListColumns("RequestDate").Index).value = requestDate
    lr.Range.Cells(1, lo.ListColumns("RequesterName").Index).value = requesterName
    lr.Range.Cells(1, lo.ListColumns("RequesterEmail").Index).value = requesterEmail
    lr.Range.Cells(1, lo.ListColumns("TxnID").Index).value = txnId
    lr.Range.Cells(1, lo.ListColumns("EventID").Index).value = eventId
    lr.Range.Cells(1, lo.ListColumns("AmountRequested").Index).value = amountRequested
    lr.Range.Cells(1, lo.ListColumns("Reason").Index).value = reason
    lr.Range.Cells(1, lo.ListColumns("Status").Index).value = "Pending"
    lr.Range.Cells(1, lo.ListColumns("CreatedAt").Index).value = Now
    lr.Range.Cells(1, lo.ListColumns("UpdatedAt").Index).value = Now

    AuditLog "RefundRequest", refundId, requesterName
    AddRefundRequest = refundId
    Exit Function
EH:
    HandleError "AddRefundRequest", Err, requesterName
    AddRefundRequest = ""
End Function

Public Sub ApproveRefundRequest(ByVal refundId As String, ByVal approvedBy As String)
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_COMPLIANCE, T_COMPLIANCE_REFUNDS)
    Dim rowIdx As Long: rowIdx = FindRowIndex(lo, "RefundID", refundId)
    If rowIdx = 0 Then Err.Raise vbObjectError + 820, "ApproveRefundRequest", "Refund not found"

    Dim requesterName As String
    requesterName = CStr(lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("RequesterName").Index).value)
    Dim allowSelf As Boolean
    allowSelf = (LCase$(GetConfigValue(CFG_REFUND_APPROVAL_SELF_OK, "FALSE")) = "true")
    If Not allowSelf And LCase$(requesterName) = LCase$(approvedBy) Then
        Err.Raise vbObjectError + 821, "ApproveRefundRequest", "Self-approval not allowed"
    End If

    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("Status").Index).value = "Approved"
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ApprovedBy").Index).value = approvedBy
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("ApprovedAt").Index).value = Now
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("UpdatedAt").Index).value = Now
    AuditLog "RefundApproved", refundId, approvedBy
    Exit Sub
EH:
    HandleError "ApproveRefundRequest", Err, refundId
End Sub

Public Function AddIncidentComplaint(ByVal caseType As String, ByVal reportDate As Date, ByVal reportedByName As String, _
                                     ByVal reportedByContact As String, ByVal severity As String, _
                                     Optional ByVal eventId As String = "", Optional ByVal summary As String = "", _
                                     Optional ByVal sscConsentFlag As String = "N") As String
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_COMPLIANCE, T_COMPLIANCE_INCIDENTS)
    Dim lr As ListRow: Set lr = lo.ListRows.Add

    Dim caseId As String
    caseId = "CASE-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(Int((9999 * Rnd) + 1), "0000")

    lr.Range.Cells(1, lo.ListColumns("CaseID").Index).value = caseId
    lr.Range.Cells(1, lo.ListColumns("CaseType").Index).value = caseType
    lr.Range.Cells(1, lo.ListColumns("ReportDate").Index).value = reportDate
    lr.Range.Cells(1, lo.ListColumns("ReportedByName").Index).value = reportedByName
    lr.Range.Cells(1, lo.ListColumns("ReportedByContact").Index).value = reportedByContact
    lr.Range.Cells(1, lo.ListColumns("EventID").Index).value = eventId
    lr.Range.Cells(1, lo.ListColumns("Summary").Index).value = summary
    lr.Range.Cells(1, lo.ListColumns("Severity").Index).value = severity
    lr.Range.Cells(1, lo.ListColumns("SSC_ConsentRelatedFlag").Index).value = sscConsentFlag
    lr.Range.Cells(1, lo.ListColumns("Status").Index).value = "Open"
    lr.Range.Cells(1, lo.ListColumns("RetentionUntil").Index).value = RetentionUntilForIncident(reportDate, severity, sscConsentFlag)
    lr.Range.Cells(1, lo.ListColumns("CreatedAt").Index).value = Now
    lr.Range.Cells(1, lo.ListColumns("UpdatedAt").Index).value = Now

    AuditLog "IncidentCase", caseId, reportedByName
    AddIncidentComplaint = caseId
    Exit Function
EH:
    HandleError "AddIncidentComplaint", Err, reportedByName
    AddIncidentComplaint = ""
End Function

Public Function AddOutcome(ByVal caseId As String, ByVal outcomeDate As Date, ByVal findings As String, _
                           ByVal consequenceType As String, Optional ByVal consequenceDetails As String = "") As String
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_COMPLIANCE, T_COMPLIANCE_OUTCOMES)
    Dim lr As ListRow: Set lr = lo.ListRows.Add

    Dim outcomeId As String
    outcomeId = "OUT-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(Int((9999 * Rnd) + 1), "0000")

    lr.Range.Cells(1, lo.ListColumns("OutcomeID").Index).value = outcomeId
    lr.Range.Cells(1, lo.ListColumns("CaseID").Index).value = caseId
    lr.Range.Cells(1, lo.ListColumns("OutcomeDate").Index).value = outcomeDate
    lr.Range.Cells(1, lo.ListColumns("Findings").Index).value = findings
    lr.Range.Cells(1, lo.ListColumns("ConsequenceType").Index).value = consequenceType
    lr.Range.Cells(1, lo.ListColumns("ConsequenceDetails").Index).value = consequenceDetails
    lr.Range.Cells(1, lo.ListColumns("CreatedAt").Index).value = Now

    LinkOutcomeToCase caseId, outcomeId
    AuditLog "OutcomeCreated", outcomeId, caseId
    AddOutcome = outcomeId
    Exit Function
EH:
    HandleError "AddOutcome", Err, caseId
    AddOutcome = ""
End Function

Public Function AddAppeal(ByVal outcomeId As String, ByVal appealDate As Date, ByVal appellantName As String, ByVal grounds As String) As String
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_COMPLIANCE, T_COMPLIANCE_APPEALS)
    Dim lr As ListRow: Set lr = lo.ListRows.Add

    Dim appealId As String
    appealId = "APL-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(Int((9999 * Rnd) + 1), "0000")

    lr.Range.Cells(1, lo.ListColumns("AppealID").Index).value = appealId
    lr.Range.Cells(1, lo.ListColumns("OutcomeID").Index).value = outcomeId
    lr.Range.Cells(1, lo.ListColumns("AppealDate").Index).value = appealDate
    lr.Range.Cells(1, lo.ListColumns("AppellantName").Index).value = appellantName
    lr.Range.Cells(1, lo.ListColumns("Grounds").Index).value = grounds
    lr.Range.Cells(1, lo.ListColumns("Status").Index).value = "Pending"
    lr.Range.Cells(1, lo.ListColumns("CreatedAt").Index).value = Now

    AuditLog "AppealCreated", appealId, outcomeId
    AddAppeal = appealId
    Exit Function
EH:
    HandleError "AddAppeal", Err, outcomeId
    AddAppeal = ""
End Function

Public Sub ExportEventInfoPdf(ByVal eventId As String)
    ExportCompliancePdf T_COMPLIANCE_EVENTS, "EventID", eventId, "Event Info"
End Sub

Public Sub ExportRefundPdf(ByVal refundId As String)
    ExportCompliancePdf T_COMPLIANCE_REFUNDS, "RefundID", refundId, "Refund Request"
End Sub

Public Sub ExportIncidentPdf(ByVal caseId As String)
    ExportCompliancePdf T_COMPLIANCE_INCIDENTS, "CaseID", caseId, "Incident/Complaint"
End Sub

Public Sub ExportGiveawayPdf(ByVal giveawayId As String)
    ExportCompliancePdf T_COMPLIANCE_GIVEAWAYS, "GiveawayID", giveawayId, "Giveaway"
End Sub

Private Sub ExportCompliancePdf(ByVal tableName As String, ByVal idColumn As String, ByVal idValue As String, ByVal titlePrefix As String)
    Dim prevSU As Boolean, prevEV As Boolean
    prevSU = Application.ScreenUpdating
    prevEV = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetTable(SH_COMPLIANCE, tableName)
    Dim rowIdx As Long: rowIdx = FindRowIndex(lo, idColumn, idValue)
    If rowIdx = 0 Then Err.Raise vbObjectError + 822, "ExportCompliancePdf", "Record not found"

    Dim labels() As String
    Dim values() As String
    Dim i As Long
    ReDim labels(1 To lo.ListColumns.count)
    ReDim values(1 To lo.ListColumns.count)
    For i = 1 To lo.ListColumns.count
        labels(i) = lo.ListColumns(i).name
        values(i) = CStr(lo.DataBodyRange.Cells(rowIdx, i).value)
    Next i

    WriteComplianceReport titlePrefix & " - " & idValue, labels, values

    Dim pdfPath As String
    pdfPath = ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_COMPLIANCE_DOCS, ".\Compliance\Docs\")) & _
        titlePrefix & "_" & idValue & ".pdf"
    GetSheet(SH_COMPLIANCE_REPORT).ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    AuditLog "ExportCompliancePdf", idValue, pdfPath
    GoTo CleanExit
EH:
    HandleError "ExportCompliancePdf", Err, idValue
CleanExit:
    Application.EnableEvents = prevEV
    Application.ScreenUpdating = prevSU
End Sub

Private Sub WriteComplianceReport(ByVal titleText As String, ByVal labels As Variant, ByVal values As Variant)
    Dim ws As Worksheet: Set ws = GetSheet(SH_COMPLIANCE_REPORT)
    ws.Cells.ClearContents
    ws.Range("A1").value = titleText
    ws.Range("A2").value = "Generated"
    ws.Range("B2").value = Now

    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        ws.Cells(3 + i, 1).value = labels(i)
        ws.Cells(3 + i, 2).value = values(i)
    Next i
    ws.Columns("A:B").ColumnWidth = 28
End Sub

Private Function RetentionUntilForIncident(ByVal reportDate As Date, ByVal severity As String, ByVal sscConsentFlag As String) As Date
    Dim years As Long
    If LCase$(severity) = "severe" Or LCase$(sscConsentFlag) = "y" Then
        years = CLng(GetConfigValue(CFG_RETENTION_SEVERE, "7"))
    Else
        years = CLng(GetConfigValue(CFG_RETENTION_MINOR, "3"))
    End If
    RetentionUntilForIncident = DateAdd("yyyy", years, reportDate)
End Function

Private Function FindRowIndex(ByVal lo As ListObject, ByVal columnName As String, ByVal value As String) As Long
    FindRowIndex = 0
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim colIdx As Long: colIdx = lo.ListColumns(columnName).Index
    Dim i As Long
    For i = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(i, colIdx).value) = value Then
            FindRowIndex = i
            Exit Function
        End If
    Next i
End Function

Private Sub LinkOutcomeToCase(ByVal caseId As String, ByVal outcomeId As String)
    Dim lo As ListObject: Set lo = GetTable(SH_COMPLIANCE, T_COMPLIANCE_INCIDENTS)
    Dim rowIdx As Long: rowIdx = FindRowIndex(lo, "CaseID", caseId)
    If rowIdx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("OutcomeID").Index).value = outcomeId
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("Status").Index).value = "Closed"
    lo.DataBodyRange.Cells(rowIdx, lo.ListColumns("UpdatedAt").Index).value = Now
End Sub

Private Function HashFile(ByVal filePath As String) As String
    On Error GoTo EH
    Dim f As Integer: f = FreeFile
    Dim fileText As String
    Open filePath For Binary As #f
    fileText = Space$(LOF(f))
    Get #f, , fileText
    Close #f
    HashFile = CStr(CLng(Crc32(fileText)))
    Exit Function
EH:
    HashFile = CStr(FileLen(filePath)) & "-" & Format$(FileDateTime(filePath), "yyyymmddhhnnss")
End Function

Private Function HashRow(ByVal rowText As String) As String
    HashRow = CStr(CLng(Crc32(rowText)))
End Function

'========================
' Self test
'========================

Private Function SelfTestReport() As String
    Dim msg As String
    msg = "Self-test results:" & vbCrLf

    msg = msg & CheckTable(SH_LEDGER, T_LEDGER) & vbCrLf
    msg = msg & CheckTable(SH_RECEIPTS, T_RECEIPTS) & vbCrLf
    msg = msg & CheckTable(SH_BUDGET, T_BUDGET) & vbCrLf
    msg = msg & CheckTable(SH_MONTHSTATUS, T_MONTHSTATUS) & vbCrLf
    msg = msg & CheckTable(SH_MEMBERS, T_MEMBERS) & vbCrLf
    msg = msg & CheckTable(SH_MEETINGS, T_MEETINGS) & vbCrLf
    msg = msg & CheckTable(SH_AGENDA, T_AGENDA) & vbCrLf
    msg = msg & CheckTable(SH_IMPORTS, T_IMPORTLOG) & vbCrLf
    msg = msg & CheckTable(SH_ERRORLOG, T_ERRORLOG) & vbCrLf
    msg = msg & CheckTable(SH_COMPLIANCE, T_COMPLIANCE_EVENTS) & vbCrLf
    msg = msg & CheckTable(SH_COMPLIANCE, T_COMPLIANCE_REFUNDS) & vbCrLf
    msg = msg & CheckTable(SH_COMPLIANCE, T_COMPLIANCE_INCIDENTS) & vbCrLf

    msg = msg & "Paths:" & vbCrLf
    msg = msg & "- Board packets: " & ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_BOARDPACKETS, ".\BoardPackets\")) & vbCrLf
    msg = msg & "- Minutes DOCX: " & ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_MINUTES_DOCX, ".\Minutes\DOCX\")) & vbCrLf
    msg = msg & "- Minutes PDF: " & ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_MINUTES_PDF, ".\Minutes\PDF\")) & vbCrLf
    msg = msg & "- Agenda DOCX: " & ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_AGENDA_DOCX, ".\Agenda\DOCX\")) & vbCrLf
    msg = msg & "- Agenda PDF: " & ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_AGENDA_PDF, ".\Agenda\PDF\")) & vbCrLf
    msg = msg & "- Imports Zeffy: " & ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_IMPORTS_ZEFFY, ".\Imports\Zeffy\")) & vbCrLf
    msg = msg & "- Imports Blaze: " & ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_IMPORTS_BLAZE, ".\Imports\Blaze\")) & vbCrLf
    msg = msg & "- Compliance Docs: " & ResolveWorkbookRelativePath(GetConfigValue(CFG_PATH_COMPLIANCE_DOCS, ".\Compliance\Docs\")) & vbCrLf
    msg = msg & "Templates:" & vbCrLf
    msg = msg & "- Minutes template: " & TemplateStatus("TCPP Board Meeting Minutes Template.docx") & vbCrLf
    msg = msg & "- Agenda template: " & TemplateStatus("Template Meeting Agenda.docx") & vbCrLf

    SelfTestReport = msg
End Function

Private Function TemplateStatus(ByVal templateName As String) As String
    Dim p As String: p = GetTemplatePath(templateName)
    If Len(Dir(p)) = 0 Then
        TemplateStatus = "MISSING (" & p & ")"
    Else
        TemplateStatus = "OK (" & p & ")"
    End If
End Function

Private Function CheckTable(ByVal sheetName As String, ByVal tableName As String) As String
    On Error GoTo EH
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    CheckTable = "OK: " & sheetName & "!" & tableName
    Exit Function
EH:
    CheckTable = "MISSING: " & sheetName & "!" & tableName
End Function

'========================
' Folder helpers
'========================

Private Function ResolveWorkbookRelativePath(ByVal rel As String) As String
    Dim p As String: p = ThisWorkbook.path
    If Len(p) = 0 Then p = CurDir$
    ResolveWorkbookRelativePath = EnsureTrailingSlash(p) & Replace(rel, ".\", "")
End Function

Private Function EnsureTrailingSlash(ByVal p As String) As String
    If right$(p, 1) = "\\" Then
        EnsureTrailingSlash = p
    Else
        EnsureTrailingSlash = p & "\\"
    End If
End Function

Private Sub EnsureFolderPath(ByVal folderPath As String)
    Dim parts() As String: parts = Split(folderPath, "\\")
    Dim i As Long, cur As String

    If InStr(folderPath, ":\\") > 0 Then
        cur = parts(0) & "\\"
        i = 1
    Else
        cur = ""
        i = 0
    End If

    For i = i To UBound(parts)
        If Len(parts(i)) > 0 Then
            cur = cur & parts(i) & "\\"
            If Dir(cur, vbDirectory) = "" Then MkDir cur
        End If
    Next i
End Sub


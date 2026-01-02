VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFixIssues 
   Caption         =   "lblEntryFrmHdr"
   ClientHeight    =   7860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12465
   OleObjectBlob   =   "frmFixIssues.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFixIssues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mIssue As String
Private mMonthKey As String

Public Sub Init(ByVal issueLabel As String, ByVal monthKey As String)
    mIssue = issueLabel
    mMonthKey = monthKey
End Sub

Private Sub UserForm_Initialize()
    cboIssueType.Clear
    cboIssueType.AddItem "Uncategorized"
    cboIssueType.AddItem "Missing Receipt"
    cboIssueType.AddItem "Not Reconciled"
    cboIssueType.AddItem "Not Closed"
    cboIssueType.value = NormalizeIssue(mIssue)

    LoadMonthList
    cboMonth.value = IIf(Len(mMonthKey) = 0, Format(Date, "yyyy-mm"), mMonthKey)

    LoadCategories
    LoadEvents
    LoadCharities

    lstIssues.ColumnCount = 6
    lstIssues.ColumnWidths = "160;60;70;160;120;90"

    RefreshIssues
End Sub

Private Sub cboIssueType_Change()
    RefreshIssues
End Sub

Private Sub cboMonth_Change()
    RefreshIssues
End Sub

Private Sub lstIssues_Click()
    LoadSelectedToEditor
End Sub

Private Sub cmdSaveChanges_Click()
    If lstIssues.ListIndex < 0 Then Exit Sub
    Dim txnId As String: txnId = CStr(lstIssues.List(lstIssues.ListIndex, 0))

    On Error Resume Next
    modTCPPv2.UpdateLedgerFields txnId, Trim$(cboCategory.value), Trim$(cboEvent.value), Trim$(cboCharity.value), CBool(chkReceiptRequired.value)
    On Error GoTo 0

    RefreshIssues
End Sub

Private Sub cmdAttachReceipt_Click()
    If lstIssues.ListIndex < 0 Then Exit Sub
    Dim txnId As String: txnId = CStr(lstIssues.List(lstIssues.ListIndex, 0))
    On Error Resume Next
    modTCPPv2.AttachReceiptToTxn txnId
    On Error GoTo 0
    RefreshIssues
End Sub

Private Sub cmdWaiveReceipt_Click()
    If lstIssues.ListIndex < 0 Then Exit Sub
    Dim txnId As String: txnId = CStr(lstIssues.List(lstIssues.ListIndex, 0))
    modTCPPv2.WaiveReceipt txnId, Trim$(txtWaiveReason.value)
    RefreshIssues
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

'----------------------
Private Sub RefreshIssues()
    lstIssues.Clear

    Dim issue As String: issue = NormalizeIssue(CStr(cboIssueType.value))
    Dim mk As String: mk = CStr(cboMonth.value)

    If issue = "Not Reconciled" Or issue = "Not Closed" Then
        ' month-level issue; no txn list
        Exit Sub
    End If

    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("DATA_Ledger").ListObjects("tblLedger")
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim i As Long
    Dim mkCol As Long: mkCol = lo.ListColumns("MonthKey").Index
    Dim idCol As Long: idCol = lo.ListColumns("TxnID").Index
    Dim dCol As Long: dCol = lo.ListColumns("Date").Index
    Dim netCol As Long: netCol = lo.ListColumns("Net").Index
    Dim payeeCol As Long: payeeCol = lo.ListColumns("PayeeOrSource").Index
    Dim catCol As Long: catCol = lo.ListColumns("Category").Index
    Dim rrCol As Long: rrCol = lo.ListColumns("ReceiptRequired").Index
    Dim rsCol As Long: rsCol = lo.ListColumns("ReceiptStatus").Index

    For i = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(i, mkCol).value) = mk Then
            Dim showRow As Boolean: showRow = False

            If issue = "Uncategorized" Then
                showRow = (Len(Trim$(CStr(lo.DataBodyRange.Cells(i, catCol).value))) = 0)
            ElseIf issue = "Missing Receipt" Then
                Dim rr As Boolean: rr = CBool(lo.DataBodyRange.Cells(i, rrCol).value)
                Dim rs As String: rs = CStr(lo.DataBodyRange.Cells(i, rsCol).value)
                showRow = (rr And rs <> "Linked" And rs <> "Waived")
            End If

            If showRow Then
                lstIssues.AddItem CStr(lo.DataBodyRange.Cells(i, idCol).value)
                lstIssues.List(lstIssues.ListCount - 1, 1) = Format(CDate(lo.DataBodyRange.Cells(i, dCol).value), "m/d")
                lstIssues.List(lstIssues.ListCount - 1, 2) = Format(CDbl(lo.DataBodyRange.Cells(i, netCol).value), "0.00")
                lstIssues.List(lstIssues.ListCount - 1, 3) = CStr(lo.DataBodyRange.Cells(i, payeeCol).value)
                lstIssues.List(lstIssues.ListCount - 1, 4) = CStr(lo.DataBodyRange.Cells(i, catCol).value)
                lstIssues.List(lstIssues.ListCount - 1, 5) = CStr(lo.DataBodyRange.Cells(i, rsCol).value)
            End If
        End If
    Next i

    If lstIssues.ListCount > 0 Then
        lstIssues.ListIndex = 0
        LoadSelectedToEditor
    End If
End Sub

Private Sub LoadSelectedToEditor()
    If lstIssues.ListIndex < 0 Then Exit Sub
    Dim txnId As String: txnId = CStr(lstIssues.List(lstIssues.ListIndex, 0))

    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("DATA_Ledger").ListObjects("tblLedger")

    Dim i As Long, idCol As Long
    idCol = lo.ListColumns("TxnID").Index

    For i = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(i, idCol).value) = txnId Then
            cboCategory.value = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Category").Index).value)
            cboEvent.value = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Event").Index).value)
            cboCharity.value = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Charity").Index).value)
            chkReceiptRequired.value = CBool(lo.DataBodyRange.Cells(i, lo.ListColumns("ReceiptRequired").Index).value)
            Exit Sub
        End If
    Next i
End Sub

Private Sub LoadMonthList()
    cboMonth.Clear
    Dim i As Long, d As Date
    d = DateSerial(Year(Date), Month(Date), 1)
    For i = -12 To 12
        cboMonth.AddItem Format(DateAdd("m", i, d), "yyyy-mm")
    Next i
End Sub

Private Sub LoadCategories()
    cboCategory.Clear
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblCOA")
    If Not lo.DataBodyRange Is Nothing Then
        Dim c As Range
        For Each c In lo.ListColumns(1).DataBodyRange.Cells
            cboCategory.AddItem CStr(c.value)
        Next c
    End If
End Sub

Private Sub LoadEvents()
    cboEvent.Clear
    cboEvent.AddItem ""
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblEvents")
    If Not lo.DataBodyRange Is Nothing Then
        Dim c As Range
        For Each c In lo.ListColumns(1).DataBodyRange.Cells
            cboEvent.AddItem CStr(c.value)
        Next c
    End If
End Sub

Private Sub LoadCharities()
    cboCharity.Clear
    cboCharity.AddItem ""
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblCharities")
    If Not lo.DataBodyRange Is Nothing Then
        Dim c As Range
        For Each c In lo.ListColumns(1).DataBodyRange.Cells
            cboCharity.AddItem CStr(c.value)
        Next c
    End If
End Sub

Private Function NormalizeIssue(ByVal s As String) As String
    If InStr(1, s, "Uncategorized", vbTextCompare) > 0 Then NormalizeIssue = "Uncategorized": Exit Function
    If InStr(1, s, "Missing Receipt", vbTextCompare) > 0 Then NormalizeIssue = "Missing Receipt": Exit Function
    If InStr(1, s, "Not Reconciled", vbTextCompare) > 0 Then NormalizeIssue = "Not Reconciled": Exit Function
    If InStr(1, s, "Not Closed", vbTextCompare) > 0 Then NormalizeIssue = "Not Closed": Exit Function
    NormalizeIssue = "Uncategorized"
End Function


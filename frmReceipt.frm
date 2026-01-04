VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReceipt 
   Caption         =   "UserForm1"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12165
   OleObjectBlob   =   "frmReceipt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMonthKey As String

Public Sub InitForMonth(ByVal monthKey As String)
    mMonthKey = monthKey
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo EH

    LoadMonthList
    If Len(mMonthKey) = 0 Then mMonthKey = Format(Date, "yyyy-mm")
    cboMonth.value = mMonthKey
    lstTxns.ColumnCount = 6
    lstTxns.ColumnWidths = "160;60;70;160;120;90"
    RefreshList
    Exit Sub
EH:
    modTCPPv2.HandleError "frmReceipt.Initialize", Err, ""
End Sub

Private Sub cboMonth_Change()
    mMonthKey = cboMonth.value
    RefreshList
End Sub

Private Sub cmdAttach_Click()
    If lstTxns.ListIndex < 0 Then Exit Sub
    On Error GoTo EH

    Dim txnId As String: txnId = CStr(lstTxns.List(lstTxns.ListIndex, 0))
    Dim vendor As String: vendor = InputBox("Vendor (optional):", "Receipt Vendor")
    Dim storage As String: storage = InputBox("Storage location/path (optional):", "Receipt Storage")
    Dim notes As String: notes = InputBox("Notes (optional):", "Receipt Notes")

    modTCPPv2.CreateReceiptInfo txnId, vendor, Date, Date, storage, notes
    RefreshList
    Exit Sub
EH:
    modTCPPv2.HandleError "frmReceipt.Attach", Err, txnId
End Sub

Private Sub cmdWaive_Click()
    If lstTxns.ListIndex < 0 Then Exit Sub
    On Error GoTo EH
    Dim txnId As String: txnId = CStr(lstTxns.List(lstTxns.ListIndex, 0))
    If Len(Trim$(txtWaiveReason.value)) = 0 Then
        MsgBox "Waive reason is required.", vbExclamation, "Receipt Waiver"
        Exit Sub
    End If
    modTCPPv2.WaiveReceipt txnId, Trim$(txtWaiveReason.value)
    RefreshList
    Exit Sub
EH:
    modTCPPv2.HandleError "frmReceipt.Waive", Err, txnId
End Sub

Private Sub cmdOpenFile_Click()
    If lstTxns.ListIndex < 0 Then Exit Sub
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

'----------------------
Private Sub RefreshList()
    lstTxns.Clear

    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("DATA_Ledger").ListObjects("tblLedger")
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim i As Long
    Dim mkCol As Long: mkCol = lo.ListColumns("MonthKey").Index
    Dim idCol As Long: idCol = lo.ListColumns("TxnID").Index
    Dim dCol As Long: dCol = lo.ListColumns("Date").Index
    Dim netCol As Long: netCol = lo.ListColumns("Net").Index
    Dim srcCol As Long: srcCol = lo.ListColumns("SourceName").Index
    Dim catCol As Long: catCol = lo.ListColumns("Category").Index
    Dim rrCol As Long: rrCol = lo.ListColumns("ReceiptRequired").Index
    Dim rsCol As Long: rsCol = lo.ListColumns("ReceiptStatus").Index
    Dim evCol As Long: evCol = lo.ListColumns("Event").Index
    Dim chCol As Long: chCol = lo.ListColumns("Charity").Index

    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, mkCol).value) = mMonthKey Then
            If modTCPPv2.gEventFilter <> "(All)" Then
                If CStr(lo.DataBodyRange.Cells(i, evCol).value) <> modTCPPv2.gEventFilter Then GoTo ContinueRow
            End If
            If modTCPPv2.gCharityFilter <> "(All)" Then
                If CStr(lo.DataBodyRange.Cells(i, chCol).value) <> modTCPPv2.gCharityFilter Then GoTo ContinueRow
            End If
            Dim rr As Boolean: rr = CBool(lo.DataBodyRange.Cells(i, rrCol).value)
            Dim rs As String: rs = CStr(lo.DataBodyRange.Cells(i, rsCol).value)
            If rr And rs <> "Recorded" And rs <> "Waived" Then
                lstTxns.AddItem CStr(lo.DataBodyRange.Cells(i, idCol).value)
                lstTxns.List(lstTxns.ListCount - 1, 1) = Format(CDate(lo.DataBodyRange.Cells(i, dCol).value), "m/d")
                lstTxns.List(lstTxns.ListCount - 1, 2) = Format(CDbl(lo.DataBodyRange.Cells(i, netCol).value), "0.00")
                lstTxns.List(lstTxns.ListCount - 1, 3) = CStr(lo.DataBodyRange.Cells(i, srcCol).value)
                lstTxns.List(lstTxns.ListCount - 1, 4) = CStr(lo.DataBodyRange.Cells(i, catCol).value)
                lstTxns.List(lstTxns.ListCount - 1, 5) = rs
            End If
        End If
ContinueRow:
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

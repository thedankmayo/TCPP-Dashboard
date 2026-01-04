VERSION 5.00
Begin VB.UserForm frmReceiptInfo
   Caption         =   "Receipt Info"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboMonth
      Height          =   285
      Left            =   120
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox txtTxnId
      Height          =   285
      Left            =   1440
      Top             =   120
      Width           =   1500
   End
   Begin VB.TextBox txtVendor
      Height          =   285
      Left            =   3120
      Top             =   120
      Width           =   1800
   End
   Begin VB.CommandButton cmdSearch
      Caption         =   "Search"
      Height          =   300
      Left            =   5040
      Top             =   120
      Width           =   900
   End
   Begin VB.ListBox lstTxns
      Height          =   4200
      Left            =   120
      Top             =   540
      Width           =   11700
   End
   Begin VB.CommandButton cmdRecord
      Caption         =   "Record Receipt"
      Height          =   360
      Left            =   120
      Top             =   4980
      Width           =   1500
   End
   Begin VB.CommandButton cmdWaive
      Caption         =   "Waive"
      Height          =   360
      Left            =   1800
      Top             =   4980
      Width           =   900
   End
   Begin VB.TextBox txtWaiveReason
      Height          =   285
      Left            =   2880
      Top             =   5040
      Width           =   2400
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   360
      Left            =   10320
      Top             =   4980
      Width           =   900
   End
End
Attribute VB_Name = "frmReceiptInfo"
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
    modTCPPv2.ApplyTheme Me
    LoadMonthList
    If Len(mMonthKey) = 0 Then mMonthKey = Format(Date, "yyyy-mm")
    cboMonth.value = mMonthKey

    lstTxns.ColumnCount = 7
    lstTxns.ColumnWidths = "140;70;80;180;140;90;120"
    RefreshList
    Exit Sub
EH:
    modTCPPv2.HandleError "frmReceiptInfo.Initialize", Err, ""
End Sub

Private Sub cmdSearch_Click()
    RefreshList
End Sub

Private Sub cmdRecord_Click()
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
    modTCPPv2.HandleError "frmReceiptInfo.Record", Err, txnId
End Sub

Private Sub cmdWaive_Click()
    If lstTxns.ListIndex < 0 Then Exit Sub
    If Len(Trim$(txtWaiveReason.value)) = 0 Then
        MsgBox "Waive reason is required.", vbExclamation, "Receipt Waiver"
        Exit Sub
    End If
    On Error GoTo EH
    Dim txnId As String: txnId = CStr(lstTxns.List(lstTxns.ListIndex, 0))
    modTCPPv2.WaiveReceipt txnId, Trim$(txtWaiveReason.value)
    RefreshList
    Exit Sub
EH:
    modTCPPv2.HandleError "frmReceiptInfo.Waive", Err, ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub RefreshList()
    lstTxns.Clear

    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("DATA_Ledger").ListObjects("tblLedger")
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim vendorByTxn As Object
    Set vendorByTxn = CreateObject("Scripting.Dictionary")
    vendorByTxn.CompareMode = 1

    Dim rLo As ListObject
    On Error Resume Next
    Set rLo = ThisWorkbook.Worksheets("DATA_Receipts").ListObjects("tblReceipts")
    On Error GoTo 0
    If Not rLo Is Nothing Then
        If Not rLo.DataBodyRange Is Nothing Then
            Dim r As Long
            For r = 1 To rLo.ListRows.Count
                Dim tId As String
                tId = CStr(rLo.DataBodyRange.Cells(r, rLo.ListColumns("TxnID").Index).Value)
                If Len(tId) > 0 Then
                    vendorByTxn(tId) = CStr(rLo.DataBodyRange.Cells(r, rLo.ListColumns("Vendor").Index).Value)
                End If
            Next r
        End If
    End If

    Dim mkFilter As String: mkFilter = Trim$(cboMonth.value)
    Dim txnFilter As String: txnFilter = Trim$(txtTxnId.value)
    Dim vendorFilter As String: vendorFilter = Trim$(txtVendor.value)

    Dim mkCol As Long: mkCol = lo.ListColumns("MonthKey").Index
    Dim idCol As Long: idCol = lo.ListColumns("TxnID").Index
    Dim dCol As Long: dCol = lo.ListColumns("Date").Index
    Dim netCol As Long: netCol = lo.ListColumns("Net").Index
    Dim srcCol As Long: srcCol = lo.ListColumns("SourceName").Index
    Dim rsCol As Long: rsCol = lo.ListColumns("ReceiptStatus").Index
    Dim rrCol As Long: rrCol = lo.ListColumns("ReceiptRequired").Index

    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If Len(mkFilter) > 0 Then
            If CStr(lo.DataBodyRange.Cells(i, mkCol).value) <> mkFilter Then GoTo ContinueRow
        End If

        Dim txnId As String: txnId = CStr(lo.DataBodyRange.Cells(i, idCol).value)
        If Len(txnFilter) > 0 Then
            If InStr(1, txnId, txnFilter, vbTextCompare) = 0 Then GoTo ContinueRow
        End If

        If CBool(lo.DataBodyRange.Cells(i, rrCol).value) Then
            Dim rs As String: rs = CStr(lo.DataBodyRange.Cells(i, rsCol).value)
            If rs <> "Recorded" And rs <> "Waived" Then
                If Len(vendorFilter) > 0 Then
                    Dim src As String: src = CStr(lo.DataBodyRange.Cells(i, srcCol).value)
                    Dim vend As String: vend = ""
                    If vendorByTxn.Exists(txnId) Then vend = CStr(vendorByTxn(txnId))
                    If InStr(1, src & " " & vend, vendorFilter, vbTextCompare) = 0 Then GoTo ContinueRow
                End If

                lstTxns.AddItem txnId
                lstTxns.List(lstTxns.ListCount - 1, 1) = Format(CDate(lo.DataBodyRange.Cells(i, dCol).value), "m/d")
                lstTxns.List(lstTxns.ListCount - 1, 2) = Format(CDbl(lo.DataBodyRange.Cells(i, netCol).value), "0.00")
                lstTxns.List(lstTxns.ListCount - 1, 3) = CStr(lo.DataBodyRange.Cells(i, srcCol).value)
                lstTxns.List(lstTxns.ListCount - 1, 4) = rs
                lstTxns.List(lstTxns.ListCount - 1, 5) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Category").Index).value)
                lstTxns.List(lstTxns.ListCount - 1, 6) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Event").Index).value)
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

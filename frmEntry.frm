VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEntry 
   Caption         =   "UserForm1"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5955
   OleObjectBlob   =   "frmEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mInitType As String
Private mInitDetail As String
Private mInitMonth As String

Public Sub Init(ByVal txnType As String, ByVal txnDetail As String, ByVal monthKey As String)
    mInitType = txnType
    mInitDetail = txnDetail
    mInitMonth = monthKey
End Sub

Private Sub UserForm_Initialize()
    lblTxnID.caption = ""
    txtDate.value = Format(Date, "m/d/yyyy")

    LoadTypes
    LoadCategories
    LoadEvents
    LoadCharities
    LoadPaymentMethods

    If Len(mInitType) > 0 Then cboTxnType.value = mInitType
    If Len(mInitMonth) > 0 Then
        ' keep date default to month
        Dim y As Long, m As Long
        y = CLng(Left$(mInitMonth, 4))
        m = CLng(Right$(mInitMonth, 2))
        txtDate.value = Format(DateSerial(y, m, 1), "m/d/yyyy")
    End If

    If LCase$(mInitType) = "income" Then
        chkReceiptRequired.value = False
    Else
        chkReceiptRequired.value = True
    End If

    txtFees.value = "0"
    txtGross.value = ""
    UpdateNetLabel
End Sub

Private Sub txtGross_Change()
    UpdateNetLabel
End Sub

Private Sub txtFees_Change()
    UpdateNetLabel
End Sub

Private Sub cboTxnType_Change()
    If LCase$(cboTxnType.value) = "income" Then
        chkReceiptRequired.value = False
    Else
        chkReceiptRequired.value = True
    End If
End Sub

Private Sub cmdSave_Click()
    SaveEntry False
End Sub

Private Sub cmdSaveAndAttach_Click()
    SaveEntry True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'----------------------
' Core save
'----------------------

Private Sub SaveEntry(ByVal attachReceiptNow As Boolean)
    On Error GoTo EH

    Dim d As Date: d = CDate(txtDate.value)
    Dim t As String: t = Trim$(cboTxnType.value)
    Dim cat As String: cat = Trim$(cboCategory.value)
    Dim ev As String: ev = Trim$(cboEvent.value)
    Dim ch As String: ch = Trim$(cboCharity.value)

    Dim gross As Double: gross = CDbl(Val(txtGross.value))
    Dim fees As Double: fees = CDbl(Val(txtFees.value))
    Dim pm As String: pm = Trim$(cboPaymentMethod.value)
    Dim payee As String: payee = Trim$(txtPayeeSource.value)
    Dim memo As String: memo = Trim$(txtMemo.value)

    Dim rr As Boolean: rr = CBool(chkReceiptRequired.value)

    If Len(t) = 0 Then Err.Raise vbObjectError + 600, "frmEntry", "TxnType required"
    If Len(cat) = 0 And LCase$(t) <> "income" Then
        ' allow blank category only temporarily, but it becomes an exception; still allow save
    End If

    Dim txnId As String
    txnId = modTCPPv2.AddLedgerEntry(d, t, mInitDetail, cat, ev, ch, gross, fees, pm, payee, memo, rr, False)

    lblTxnID.caption = txnId

    If attachReceiptNow And rr Then
        modTCPPv2.AttachReceiptToTxn txnId
    End If

    Unload Me
    Exit Sub

EH:
    lblTxnID.caption = "ERROR: " & Err.Description
End Sub

Private Sub UpdateNetLabel()
    Dim g As Double: g = CDbl(Val(txtGross.value))
    Dim f As Double: f = CDbl(Val(txtFees.value))
    lblNet.caption = "Net: $" & Format$(g - f, "0.00")
End Sub

'----------------------
' Load lists
'----------------------

Private Sub LoadTypes()
    cboTxnType.Clear
    cboTxnType.AddItem "Income"
    cboTxnType.AddItem "Expense"
    cboTxnType.AddItem "Reimbursement"
    cboTxnType.AddItem "Adjustment"
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

Private Sub LoadPaymentMethods()
    cboPaymentMethod.Clear
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblPaymentMethods")
    If Not lo.DataBodyRange Is Nothing Then
        Dim c As Range
        For Each c In lo.ListColumns(1).DataBodyRange.Cells
            cboPaymentMethod.AddItem CStr(c.value)
        Next c
    End If
End Sub


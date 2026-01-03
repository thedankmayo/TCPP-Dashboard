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
    On Error GoTo EH

    lblTxnID.caption = ""
    txtDate.value = Format(Date, "m/d/yyyy")

    LoadTypes
    LoadCategories
    LoadEvents
    LoadCharities
    LoadPaymentMethods

    If Len(mInitType) > 0 Then cboTxnType.value = mInitType
    If Len(mInitMonth) > 0 Then
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
    Exit Sub
EH:
    modTCPPv2.HandleError "frmEntry.Initialize", Err, ""
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
    Dim sourceName As String: sourceName = Trim$(txtPayeeSource.value)
    Dim sourceType As String: sourceType = GetControlText("cboSourceType", "Other")
    Dim memberName As String: memberName = GetControlText("txtMemberName", "")
    Dim memberEmail As String: memberEmail = GetControlText("txtMemberEmail", "")
    Dim memo As String: memo = Trim$(txtMemo.value)

    Dim rr As Boolean: rr = CBool(chkReceiptRequired.value)

    If Len(t) = 0 Then Err.Raise vbObjectError + 600, "frmEntry", "TxnType required"

    Dim txnId As String
    txnId = modTCPPv2.AddLedgerEntry(d, t, mInitDetail, cat, ev, ch, gross, fees, pm, sourceType, sourceName, _
                                     memberName, memberEmail, memo, rr)

    lblTxnID.caption = txnId

    If attachReceiptNow And rr Then
        frmReceipt.InitForMonth modTCPPv2.MonthKeyFromDate(d)
        frmReceipt.Show vbModal
    End If

    Unload Me
    Exit Sub

EH:
    modTCPPv2.HandleError "frmEntry.SaveEntry", Err, ""
End Sub

Private Function GetControlText(ByVal controlName As String, ByVal fallback As String) As String
    On Error GoTo EH
    Dim ctl As MSForms.Control
    Set ctl = Me.Controls(controlName)
    GetControlText = Trim$(CStr(ctl.Value))
    If Len(GetControlText) = 0 Then GetControlText = fallback
    Exit Function
EH:
    GetControlText = fallback
End Function

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
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblTxnTypes")
    If Not lo.DataBodyRange Is Nothing Then
        Dim c As Range
        For Each c In lo.ListColumns(1).DataBodyRange.Cells
            cboTxnType.AddItem CStr(c.value)
        Next c
    End If
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

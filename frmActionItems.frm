VERSION 5.00
Begin VB.UserForm frmActionItems
   Caption         =   "Action Items"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8000
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtActionItem
      Height          =   285
      Left            =   120
      Top             =   120
      Width           =   2400
   End
   Begin VB.TextBox txtOwner
      Height          =   285
      Left            =   2640
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox txtDueDate
      Height          =   285
      Left            =   3960
      Top             =   120
      Width           =   1200
   End
   Begin VB.ComboBox cboStatus
      Height          =   285
      Left            =   5280
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdAdd
      Caption         =   "Add"
      Height          =   300
      Left            =   6600
      Top             =   120
      Width           =   720
   End
   Begin VB.CommandButton cmdDelete
      Caption         =   "Delete"
      Height          =   300
      Left            =   7320
      Top             =   120
      Width           =   720
   End
   Begin VB.ListBox lstActions
      Height          =   3300
      Left            =   120
      Top             =   540
      Width           =   7800
   End
End
Attribute VB_Name = "frmActionItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMeetingId As String

Public Sub Init(ByVal meetingId As String)
    mMeetingId = meetingId
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo EH
    cboStatus.Clear
    cboStatus.AddItem "Open"
    cboStatus.AddItem "Done"
    cboStatus.AddItem "Deferred"
    cboStatus.value = "Open"

    lstActions.ColumnCount = 4
    lstActions.ColumnWidths = "240;120;120;80"
    RefreshList
    Exit Sub
EH:
    modTCPPv2.HandleError "frmActionItems.Initialize", Err, mMeetingId
End Sub

Private Sub cmdAdd_Click()
    If Len(Trim$(txtActionItem.value)) = 0 Then Exit Sub
    On Error GoTo EH
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_ActionItems").ListObjects("tblActionItems")
    Dim lr As ListRow: Set lr = lo.ListRows.Add

    lr.Range.Cells(1, lo.ListColumns("ActionID").Index).Value = "ACT-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(Int((9999 * Rnd) + 1), "0000")
    lr.Range.Cells(1, lo.ListColumns("MeetingID").Index).Value = mMeetingId
    lr.Range.Cells(1, lo.ListColumns("ActionItem").Index).Value = Trim$(txtActionItem.value)
    lr.Range.Cells(1, lo.ListColumns("Owner").Index).Value = Trim$(txtOwner.value)
    If IsDate(txtDueDate.value) Then lr.Range.Cells(1, lo.ListColumns("DueDate").Index).Value = CDate(txtDueDate.value)
    lr.Range.Cells(1, lo.ListColumns("Status").Index).Value = cboStatus.value
    lr.Range.Cells(1, lo.ListColumns("Notes").Index).Value = ""

    RefreshList
    Exit Sub
EH:
    modTCPPv2.HandleError "frmActionItems.Add", Err, mMeetingId
End Sub

Private Sub cmdDelete_Click()
    If lstActions.ListIndex < 0 Then Exit Sub
    On Error GoTo EH
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_ActionItems").ListObjects("tblActionItems")
    Dim i As Long
    For i = lo.ListRows.Count To 1 Step -1
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingID").Index).Value) = mMeetingId And _
           CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("ActionItem").Index).Value) = CStr(lstActions.List(lstActions.ListIndex, 0)) Then
            lo.ListRows(i).Delete
            Exit For
        End If
    Next i
    RefreshList
    Exit Sub
EH:
    modTCPPv2.HandleError "frmActionItems.Delete", Err, mMeetingId
End Sub

Private Sub RefreshList()
    lstActions.Clear
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_ActionItems").ListObjects("tblActionItems")
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingID").Index).Value) = mMeetingId Then
            lstActions.AddItem CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("ActionItem").Index).Value)
            lstActions.List(lstActions.ListCount - 1, 1) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Owner").Index).Value)
            lstActions.List(lstActions.ListCount - 1, 2) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("DueDate").Index).Value)
            lstActions.List(lstActions.ListCount - 1, 3) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Status").Index).Value)
        End If
    Next i
End Sub

VERSION 5.00
Begin VB.UserForm frmAgendaLines
   Caption         =   "Agenda Lines"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8000
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLineTime
      Height          =   285
      Left            =   120
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox txtTopic
      Height          =   285
      Left            =   1440
      Top             =   120
      Width           =   1800
   End
   Begin VB.TextBox txtOwner
      Height          =   285
      Left            =   3360
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox txtActionItem
      Height          =   285
      Left            =   4680
      Top             =   120
      Width           =   1800
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
   Begin VB.ListBox lstLines
      Height          =   3300
      Left            =   120
      Top             =   540
      Width           =   7800
   End
End
Attribute VB_Name = "frmAgendaLines"
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
    lstLines.ColumnCount = 4
    lstLines.ColumnWidths = "100;200;200;200"
    RefreshList
    Exit Sub
EH:
    modTCPPv2.HandleError "frmAgendaLines.Initialize", Err, mMeetingId
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo EH
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_MinutesAgenda").ListObjects("tblMinutesAgendaLines")
    Dim lr As ListRow: Set lr = lo.ListRows.Add

    lr.Range.Cells(1, lo.ListColumns("MeetingID").Index).Value = mMeetingId
    lr.Range.Cells(1, lo.ListColumns("LineTime").Index).Value = Trim$(txtLineTime.value)
    lr.Range.Cells(1, lo.ListColumns("Topic").Index).Value = Trim$(txtTopic.value)
    lr.Range.Cells(1, lo.ListColumns("ActionItem").Index).Value = Trim$(txtActionItem.value)
    lr.Range.Cells(1, lo.ListColumns("Owner").Index).Value = Trim$(txtOwner.value)

    RefreshList
    Exit Sub
EH:
    modTCPPv2.HandleError "frmAgendaLines.Add", Err, mMeetingId
End Sub

Private Sub cmdDelete_Click()
    If lstLines.ListIndex < 0 Then Exit Sub
    On Error GoTo EH
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_MinutesAgenda").ListObjects("tblMinutesAgendaLines")
    Dim i As Long
    For i = lo.ListRows.Count To 1 Step -1
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingID").Index).Value) = mMeetingId And _
           CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Topic").Index).Value) = CStr(lstLines.List(lstLines.ListIndex, 1)) Then
            lo.ListRows(i).Delete
            Exit For
        End If
    Next i
    RefreshList
    Exit Sub
EH:
    modTCPPv2.HandleError "frmAgendaLines.Delete", Err, mMeetingId
End Sub

Private Sub RefreshList()
    lstLines.Clear
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_MinutesAgenda").ListObjects("tblMinutesAgendaLines")
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingID").Index).Value) = mMeetingId Then
            lstLines.AddItem CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("LineTime").Index).Value)
            lstLines.List(lstLines.ListCount - 1, 1) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Topic").Index).Value)
            lstLines.List(lstLines.ListCount - 1, 2) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Owner").Index).Value)
            lstLines.List(lstLines.ListCount - 1, 3) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("ActionItem").Index).Value)
        End If
    Next i
End Sub

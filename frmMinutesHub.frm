VERSION 5.00
Begin VB.UserForm frmMinutesHub
   Caption         =   "Minutes Hub"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNewMeeting
      Caption         =   "New Meeting"
      Height          =   360
      Left            =   120
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdExportPdf
      Caption         =   "Export PDF"
      Height          =   360
      Left            =   1440
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdOpenDoc
      Caption         =   "Open DOCX"
      Height          =   360
      Left            =   2760
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdOpenPdf
      Caption         =   "Open PDF"
      Height          =   360
      Left            =   4080
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdAttendance
      Caption         =   "Attendance"
      Height          =   360
      Left            =   120
      Top             =   480
      Width           =   1200
   End
   Begin VB.CommandButton cmdAgendaLines
      Caption         =   "Agenda Lines"
      Height          =   360
      Left            =   1440
      Top             =   480
      Width           =   1200
   End
   Begin VB.CommandButton cmdActionItems
      Caption         =   "Action Items"
      Height          =   360
      Left            =   2760
      Top             =   480
      Width           =   1200
   End
   Begin VB.TextBox txtSearch
      Height          =   285
      Left            =   4080
      Top             =   540
      Width           =   1800
   End
   Begin VB.CommandButton cmdSearch
      Caption         =   "Search"
      Height          =   300
      Left            =   6000
      Top             =   540
      Width           =   720
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   360
      Left            =   5520
      Top             =   120
      Width           =   1200
   End
   Begin VB.ListBox lstMeetings
      Height          =   3240
      Left            =   120
      Top             =   960
      Width           =   6960
   End
End
Attribute VB_Name = "frmMinutesHub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    On Error GoTo EH
    lstMeetings.ColumnCount = 4
    lstMeetings.ColumnWidths = "120;120;280;240"
    RefreshMeetings
    Exit Sub
EH:
    modTCPPv2.HandleError "frmMinutesHub.Initialize", Err, ""
End Sub

Private Sub cmdNewMeeting_Click()
    On Error GoTo EH
    Dim meetingDate As Date
    meetingDate = CDate(InputBox("Meeting date (YYYY-MM-DD):", "New Meeting", Format(Date, "yyyy-mm-dd")))
    Dim scribe As String
    scribe = InputBox("Scribe:", "New Meeting")
    Dim location As String
    location = InputBox("Location:", "New Meeting")

    modTCPPv2.CreateMeeting meetingDate, scribe, location
    RefreshMeetings
    Exit Sub
EH:
    modTCPPv2.HandleError "frmMinutesHub.NewMeeting", Err, ""
End Sub

Private Sub cmdExportPdf_Click()
    If lstMeetings.ListIndex < 0 Then Exit Sub
    On Error GoTo EH
    modTCPPv2.ExportMeetingPdf CStr(lstMeetings.List(lstMeetings.ListIndex, 0))
    RefreshMeetings
    Exit Sub
EH:
    modTCPPv2.HandleError "frmMinutesHub.ExportPdf", Err, ""
End Sub

Private Sub cmdOpenDoc_Click()
    If lstMeetings.ListIndex < 0 Then Exit Sub
    Dim path As String
    path = CStr(lstMeetings.List(lstMeetings.ListIndex, 2))
    If Len(path) > 0 Then ThisWorkbook.FollowHyperlink path
End Sub

Private Sub cmdOpenPdf_Click()
    If lstMeetings.ListIndex < 0 Then Exit Sub
    Dim path As String
    path = CStr(lstMeetings.List(lstMeetings.ListIndex, 3))
    If Len(path) > 0 Then ThisWorkbook.FollowHyperlink path
End Sub

Private Sub cmdAttendance_Click()
    If lstMeetings.ListIndex < 0 Then Exit Sub
    frmAttendance.Init CStr(lstMeetings.List(lstMeetings.ListIndex, 0))
    frmAttendance.Show vbModal
End Sub

Private Sub cmdAgendaLines_Click()
    If lstMeetings.ListIndex < 0 Then Exit Sub
    frmAgendaLines.Init CStr(lstMeetings.List(lstMeetings.ListIndex, 0))
    frmAgendaLines.Show vbModal
End Sub

Private Sub cmdActionItems_Click()
    If lstMeetings.ListIndex < 0 Then Exit Sub
    frmActionItems.Init CStr(lstMeetings.List(lstMeetings.ListIndex, 0))
    frmActionItems.Show vbModal
End Sub

Private Sub cmdSearch_Click()
    RefreshMeetings Trim$(txtSearch.value)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub RefreshMeetings(Optional ByVal filterText As String = "")
    lstMeetings.Clear

    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Meetings").ListObjects("tblMeetings")
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To lo.ListRows.Count
        Dim meetingId As String: meetingId = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingID").Index).Value)
        Dim scribe As String: scribe = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Scribe").Index).Value)
        Dim dateStr As String: dateStr = Format$(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingDate").Index).Value, "yyyy-mm-dd")
        If Len(filterText) = 0 Or InStr(1, LCase$(meetingId & " " & scribe & " " & dateStr), LCase$(filterText)) > 0 Then
            lstMeetings.AddItem meetingId
            lstMeetings.List(lstMeetings.ListCount - 1, 1) = dateStr
            lstMeetings.List(lstMeetings.ListCount - 1, 2) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MinutesDocPath").Index).Value)
            lstMeetings.List(lstMeetings.ListCount - 1, 3) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MinutesPdfPath").Index).Value)
        End If
    Next i
End Sub

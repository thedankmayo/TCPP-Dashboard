VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMinutesHub 
   Caption         =   "UserForm1"
   ClientHeight    =   5000
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11820
   OleObjectBlob   =   "frmMinutesHub.frx":0000
   StartUpPosition =   1  'CenterOwner
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
    For i = 1 To lo.ListRows.count
        Dim meetingId As String: meetingId = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingID").Index).value)
        Dim scribe As String: scribe = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Scribe").Index).value)
        Dim dateStr As String: dateStr = Format$(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingDate").Index).value, "yyyy-mm-dd")
        If Len(filterText) = 0 Or InStr(1, LCase$(meetingId & " " & scribe & " " & dateStr), LCase$(filterText)) > 0 Then
            lstMeetings.AddItem meetingId
            lstMeetings.List(lstMeetings.ListCount - 1, 1) = dateStr
            lstMeetings.List(lstMeetings.ListCount - 1, 2) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MinutesDocPath").Index).value)
            lstMeetings.List(lstMeetings.ListCount - 1, 3) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MinutesPdfPath").Index).value)
        End If
    Next i
End Sub

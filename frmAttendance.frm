VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAttendance 
   Caption         =   "UserForm1"
   ClientHeight    =   3810
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9720
   OleObjectBlob   =   "frmAttendance.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAttendance"
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
    lstAttendance.ColumnCount = 3
    lstAttendance.ColumnWidths = "200;160;80"
    RefreshList
    Exit Sub
EH:
    modTCPPv2.HandleError "frmAttendance.Initialize", Err, mMeetingId
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo EH
    If Len(Trim$(txtPersonName.value)) = 0 Then Exit Sub
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Attendance").ListObjects("tblAttendance")
    Dim lr As ListRow: Set lr = lo.ListRows.Add

    lr.Range.Cells(1, lo.ListColumns("MeetingID").Index).value = mMeetingId
    lr.Range.Cells(1, lo.ListColumns("PersonName").Index).value = Trim$(txtPersonName.value)
    lr.Range.Cells(1, lo.ListColumns("Role").Index).value = Trim$(txtRole.value)
    lr.Range.Cells(1, lo.ListColumns("PresentFlag").Index).value = CBool(chkPresent.value)

    RefreshList
    Exit Sub
EH:
    modTCPPv2.HandleError "frmAttendance.Add", Err, mMeetingId
End Sub

Private Sub cmdDelete_Click()
    If lstAttendance.ListIndex < 0 Then Exit Sub
    On Error GoTo EH
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Attendance").ListObjects("tblAttendance")
    Dim i As Long
    For i = lo.ListRows.count To 1 Step -1
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingID").Index).value) = mMeetingId And _
           CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("PersonName").Index).value) = CStr(lstAttendance.List(lstAttendance.ListIndex, 0)) Then
            lo.ListRows(i).Delete
            Exit For
        End If
    Next i
    RefreshList
    Exit Sub
EH:
    modTCPPv2.HandleError "frmAttendance.Delete", Err, mMeetingId
End Sub

Private Sub RefreshList()
    lstAttendance.Clear
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Attendance").ListObjects("tblAttendance")
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingID").Index).value) = mMeetingId Then
            lstAttendance.AddItem CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("PersonName").Index).value)
            lstAttendance.List(lstAttendance.ListCount - 1, 1) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Role").Index).value)
            lstAttendance.List(lstAttendance.ListCount - 1, 2) = IIf(CBool(lo.DataBodyRange.Cells(i, lo.ListColumns("PresentFlag").Index).value), "Y", "N")
        End If
    Next i
End Sub


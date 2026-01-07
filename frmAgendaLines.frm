VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAgendaLines 
   Caption         =   "UserForm1"
   ClientHeight    =   3810
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12885
   OleObjectBlob   =   "frmAgendaLines.frx":0000
   StartUpPosition =   1  'CenterOwner
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

    lr.Range.Cells(1, lo.ListColumns("MeetingID").Index).value = mMeetingId
    lr.Range.Cells(1, lo.ListColumns("LineTime").Index).value = Trim$(txtLineTime.value)
    lr.Range.Cells(1, lo.ListColumns("Topic").Index).value = Trim$(txtTopic.value)
    lr.Range.Cells(1, lo.ListColumns("ActionItem").Index).value = Trim$(txtActionItem.value)
    lr.Range.Cells(1, lo.ListColumns("Owner").Index).value = Trim$(txtOwner.value)

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
    For i = lo.ListRows.count To 1 Step -1
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingID").Index).value) = mMeetingId And _
           CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Topic").Index).value) = CStr(lstLines.List(lstLines.ListIndex, 1)) Then
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
    For i = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingID").Index).value) = mMeetingId Then
            lstLines.AddItem CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("LineTime").Index).value)
            lstLines.List(lstLines.ListCount - 1, 1) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Topic").Index).value)
            lstLines.List(lstLines.ListCount - 1, 2) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Owner").Index).value)
            lstLines.List(lstLines.ListCount - 1, 3) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("ActionItem").Index).value)
        End If
    Next i
End Sub


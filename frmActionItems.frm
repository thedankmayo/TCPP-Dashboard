VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmActionItems 
   Caption         =   "UserForm1"
   ClientHeight    =   3810
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12375
   OleObjectBlob   =   "frmActionItems.frx":0000
   StartUpPosition =   1  'CenterOwner
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

    lr.Range.Cells(1, lo.ListColumns("ActionID").Index).value = "ACT-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(Int((9999 * Rnd) + 1), "0000")
    lr.Range.Cells(1, lo.ListColumns("MeetingID").Index).value = mMeetingId
    lr.Range.Cells(1, lo.ListColumns("ActionItem").Index).value = Trim$(txtActionItem.value)
    lr.Range.Cells(1, lo.ListColumns("Owner").Index).value = Trim$(txtOwner.value)
    If IsDate(txtDueDate.value) Then lr.Range.Cells(1, lo.ListColumns("DueDate").Index).value = CDate(txtDueDate.value)
    lr.Range.Cells(1, lo.ListColumns("Status").Index).value = cboStatus.value
    lr.Range.Cells(1, lo.ListColumns("Notes").Index).value = ""

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
    For i = lo.ListRows.count To 1 Step -1
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingID").Index).value) = mMeetingId And _
           CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("ActionItem").Index).value) = CStr(lstActions.List(lstActions.ListIndex, 0)) Then
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
    For i = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingID").Index).value) = mMeetingId Then
            lstActions.AddItem CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("ActionItem").Index).value)
            lstActions.List(lstActions.ListCount - 1, 1) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Owner").Index).value)
            lstActions.List(lstActions.ListCount - 1, 2) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("DueDate").Index).value)
            lstActions.List(lstActions.ListCount - 1, 3) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Status").Index).value)
        End If
    Next i
End Sub

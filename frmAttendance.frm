VERSION 5.00
Begin VB.UserForm frmAttendance
   Caption         =   "Attendance"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPersonName
      Height          =   285
      Left            =   120
      Top             =   120
      Width           =   2000
   End
   Begin VB.TextBox txtRole
      Height          =   285
      Left            =   2280
      Top             =   120
      Width           =   1600
   End
   Begin VB.CheckBox chkPresent
      Caption         =   "Present"
      Height          =   240
      Left            =   3960
      Top             =   120
      Width           =   960
   End
   Begin VB.CommandButton cmdAdd
      Caption         =   "Add"
      Height          =   300
      Left            =   5040
      Top             =   120
      Width           =   720
   End
   Begin VB.CommandButton cmdDelete
      Caption         =   "Delete"
      Height          =   300
      Left            =   5880
      Top             =   120
      Width           =   720
   End
   Begin VB.ListBox lstAttendance
      Height          =   3000
      Left            =   120
      Top             =   540
      Width           =   6960
   End
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

    lr.Range.Cells(1, lo.ListColumns("MeetingID").Index).Value = mMeetingId
    lr.Range.Cells(1, lo.ListColumns("PersonName").Index).Value = Trim$(txtPersonName.value)
    lr.Range.Cells(1, lo.ListColumns("Role").Index).Value = Trim$(txtRole.value)
    lr.Range.Cells(1, lo.ListColumns("PresentFlag").Index).Value = CBool(chkPresent.value)

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
    For i = lo.ListRows.Count To 1 Step -1
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingID").Index).Value) = mMeetingId And _
           CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("PersonName").Index).Value) = CStr(lstAttendance.List(lstAttendance.ListIndex, 0)) Then
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
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MeetingID").Index).Value) = mMeetingId Then
            lstAttendance.AddItem CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("PersonName").Index).Value)
            lstAttendance.List(lstAttendance.ListCount - 1, 1) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Role").Index).Value)
            lstAttendance.List(lstAttendance.ListCount - 1, 2) = IIf(CBool(lo.DataBodyRange.Cells(i, lo.ListColumns("PresentFlag").Index).Value), "Y", "N")
        End If
    Next i
End Sub

VERSION 5.00
Begin VB.UserForm frmMembers
   Caption         =   "Members"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtName
      Height          =   285
      Left            =   120
      Top             =   120
      Width           =   2000
   End
   Begin VB.TextBox txtEmail
      Height          =   285
      Left            =   2280
      Top             =   120
      Width           =   2200
   End
   Begin VB.ComboBox cboMembershipType
      Height          =   285
      Left            =   4560
      Top             =   120
      Width           =   1800
   End
   Begin VB.CheckBox chkDuesPaid
      Caption         =   "Dues Paid"
      Height          =   240
      Left            =   6480
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox txtPaidDate
      Height          =   285
      Left            =   120
      Top             =   480
      Width           =   2000
   End
   Begin VB.TextBox txtDuesAmount
      Height          =   285
      Left            =   2280
      Top             =   480
      Width           =   1200
   End
   Begin VB.TextBox txtJoinedDate
      Height          =   285
      Left            =   3600
      Top             =   480
      Width           =   1200
   End
   Begin VB.TextBox txtNotes
      Height          =   285
      Left            =   4920
      Top             =   480
      Width           =   2760
   End
   Begin VB.TextBox txtRenewalDate
      Height          =   285
      Left            =   120
      Top             =   780
      Width           =   2000
   End
   Begin VB.CommandButton cmdSave
      Caption         =   "Save"
      Height          =   360
      Left            =   7800
      Top             =   120
      Width           =   960
   End
   Begin VB.CommandButton cmdSearch
      Caption         =   "Search"
      Height          =   360
      Left            =   7800
      Top             =   540
      Width           =   960
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   360
      Left            =   7800
      Top             =   960
      Width           =   960
   End
   Begin VB.CommandButton cmdExportReport
      Caption         =   "Export Report"
      Height          =   360
      Left            =   6600
      Top             =   960
      Width           =   1080
   End
   Begin VB.ListBox lstMembers
      Height          =   3600
      Left            =   120
      Top             =   960
      Width           =   7560
   End
End
Attribute VB_Name = "frmMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    On Error GoTo EH
    LoadMembershipTypes
    lstMembers.ColumnCount = 5
    lstMembers.ColumnWidths = "200;220;120;80;120"
    RefreshMembers ""
    Exit Sub
EH:
    modTCPPv2.HandleError "frmMembers.Initialize", Err, ""
End Sub

Private Sub cmdSave_Click()
    On Error GoTo EH

    modTCPPv2.UpsertMember Trim$(txtName.value), Trim$(txtEmail.value), cboMembershipType.value, _
        CBool(chkDuesPaid.value), txtPaidDate.value, CDbl(Val(txtDuesAmount.value)), txtJoinedDate.value, Trim$(txtNotes.value)

    txtRenewalDate.value = CStr(modTCPPv2.CalculateRenewalDate(txtPaidDate.value))
    RefreshMembers ""
    Exit Sub
EH:
    modTCPPv2.HandleError "frmMembers.Save", Err, ""
End Sub

Private Sub cmdSearch_Click()
    RefreshMembers Trim$(txtName.value)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExportReport_Click()
    On Error GoTo EH
    modTCPPv2.ExportDuesReport
    Exit Sub
EH:
    modTCPPv2.HandleError "frmMembers.ExportReport", Err, ""
End Sub

Private Sub lstMembers_Click()
    If lstMembers.ListIndex < 0 Then Exit Sub
    txtName.value = CStr(lstMembers.List(lstMembers.ListIndex, 0))
    txtEmail.value = CStr(lstMembers.List(lstMembers.ListIndex, 1))
    cboMembershipType.value = CStr(lstMembers.List(lstMembers.ListIndex, 2))
    chkDuesPaid.value = (CStr(lstMembers.List(lstMembers.ListIndex, 3)) = "Y")
    txtRenewalDate.value = CStr(lstMembers.List(lstMembers.ListIndex, 4))

    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Members").ListObjects("tblMembers")
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MemberEmail").Index).Value) = txtEmail.value Then
            txtPaidDate.value = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("DuesPaidDate").Index).Value)
            txtDuesAmount.value = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("DuesAmount").Index).Value)
            txtJoinedDate.value = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("JoinedDate").Index).Value)
            txtNotes.value = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Notes").Index).Value)
            Exit For
        End If
    Next i
End Sub

Private Sub LoadMembershipTypes()
    cboMembershipType.Clear
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblMembershipTypes")
    If Not lo.DataBodyRange Is Nothing Then
        Dim c As Range
        For Each c In lo.ListColumns(1).DataBodyRange.Cells
            cboMembershipType.AddItem CStr(c.value)
        Next c
    End If
End Sub

Private Sub RefreshMembers(ByVal filterName As String)
    lstMembers.Clear
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Members").ListObjects("tblMembers")
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To lo.ListRows.Count
        Dim name As String
        name = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MemberName").Index).Value)
        Dim email As String
        email = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MemberEmail").Index).Value)
        If Len(filterName) = 0 Or InStr(1, LCase$(name & " " & email), LCase$(filterName)) > 0 Then
            lstMembers.AddItem name
            lstMembers.List(lstMembers.ListCount - 1, 1) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MemberEmail").Index).Value)
            lstMembers.List(lstMembers.ListCount - 1, 2) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MembershipType").Index).Value)
            lstMembers.List(lstMembers.ListCount - 1, 3) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("DuesPaidFlag").Index).Value)
            lstMembers.List(lstMembers.ListCount - 1, 4) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("RenewalDate").Index).Value)
        End If
    Next i
End Sub

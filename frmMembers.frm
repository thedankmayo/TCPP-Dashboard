VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMembers 
   Caption         =   "UserForm1"
   ClientHeight    =   5200
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   13680
   OleObjectBlob   =   "frmMembers.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    On Error GoTo EH
    modTCPPv2.ApplyTheme Me
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
    For i = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MemberEmail").Index).value) = txtEmail.value Then
            txtPaidDate.value = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("DuesPaidDate").Index).value)
            txtDuesAmount.value = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("DuesAmount").Index).value)
            txtJoinedDate.value = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("JoinedDate").Index).value)
            txtNotes.value = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Notes").Index).value)
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
    For i = 1 To lo.ListRows.count
        Dim name As String
        name = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MemberName").Index).value)
        Dim email As String
        email = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MemberEmail").Index).value)
        If Len(filterName) = 0 Or InStr(1, LCase$(name & " " & email), LCase$(filterName)) > 0 Then
            lstMembers.AddItem name
            lstMembers.List(lstMembers.ListCount - 1, 1) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MemberEmail").Index).value)
            lstMembers.List(lstMembers.ListCount - 1, 2) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("MembershipType").Index).value)
            lstMembers.List(lstMembers.ListCount - 1, 3) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("DuesPaidFlag").Index).value)
            lstMembers.List(lstMembers.ListCount - 1, 4) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("RenewalDate").Index).value)
        End If
    Next i
End Sub


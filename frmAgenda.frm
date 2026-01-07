VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAgenda 
   Caption         =   "UserForm1"
   ClientHeight    =   4240
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11640
   OleObjectBlob   =   "frmAgenda.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    On Error GoTo EH
    lstAgendas.ColumnCount = 4
    lstAgendas.ColumnWidths = "120;120;280;240"
    RefreshAgendas
    Exit Sub
EH:
    modTCPPv2.HandleError "frmAgenda.Initialize", Err, ""
End Sub

Private Sub cmdNewAgenda_Click()
    On Error GoTo EH
    Dim agendaDate As Date
    agendaDate = CDate(InputBox("Agenda date (YYYY-MM-DD):", "New Agenda", Format(Date, "yyyy-mm-dd")))
    modTCPPv2.CreateAgenda agendaDate
    RefreshAgendas
    Exit Sub
EH:
    modTCPPv2.HandleError "frmAgenda.NewAgenda", Err, ""
End Sub

Private Sub cmdExportPdf_Click()
    If lstAgendas.ListIndex < 0 Then Exit Sub
    On Error GoTo EH
    modTCPPv2.ExportAgendaPdf CStr(lstAgendas.List(lstAgendas.ListIndex, 0))
    RefreshAgendas
    Exit Sub
EH:
    modTCPPv2.HandleError "frmAgenda.ExportPdf", Err, ""
End Sub

Private Sub cmdOpenDoc_Click()
    If lstAgendas.ListIndex < 0 Then Exit Sub
    Dim path As String
    path = CStr(lstAgendas.List(lstAgendas.ListIndex, 2))
    If Len(path) > 0 Then ThisWorkbook.FollowHyperlink path
End Sub

Private Sub cmdOpenPdf_Click()
    If lstAgendas.ListIndex < 0 Then Exit Sub
    Dim path As String
    path = CStr(lstAgendas.List(lstAgendas.ListIndex, 3))
    If Len(path) > 0 Then ThisWorkbook.FollowHyperlink path
End Sub

Private Sub cmdSearch_Click()
    RefreshAgendas Trim$(txtSearch.value)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub RefreshAgendas(Optional ByVal filterText As String = "")
    lstAgendas.Clear

    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Agenda").ListObjects("tblAgenda")
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To lo.ListRows.count
        Dim agendaId As String: agendaId = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("AgendaID").Index).value)
        Dim dateStr As String: dateStr = Format$(lo.DataBodyRange.Cells(i, lo.ListColumns("AgendaDate").Index).value, "yyyy-mm-dd")
        If Len(filterText) = 0 Or InStr(1, LCase$(agendaId & " " & dateStr), LCase$(filterText)) > 0 Then
            lstAgendas.AddItem agendaId
            lstAgendas.List(lstAgendas.ListCount - 1, 1) = dateStr
            lstAgendas.List(lstAgendas.ListCount - 1, 2) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("DocPath").Index).value)
            lstAgendas.List(lstAgendas.ListCount - 1, 3) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("PdfPath").Index).value)
        End If
    Next i
End Sub

VERSION 5.00
Begin VB.UserForm frmAgenda
   Caption         =   "Agenda Hub"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNewAgenda
      Caption         =   "New Agenda"
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
   Begin VB.ListBox lstAgendas
      Height          =   3240
      Left            =   120
      Top             =   960
      Width           =   6960
   End
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
    For i = 1 To lo.ListRows.Count
        Dim agendaId As String: agendaId = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("AgendaID").Index).Value)
        Dim dateStr As String: dateStr = Format$(lo.DataBodyRange.Cells(i, lo.ListColumns("AgendaDate").Index).Value, "yyyy-mm-dd")
        If Len(filterText) = 0 Or InStr(1, LCase$(agendaId & " " & dateStr), LCase$(filterText)) > 0 Then
            lstAgendas.AddItem agendaId
            lstAgendas.List(lstAgendas.ListCount - 1, 1) = dateStr
            lstAgendas.List(lstAgendas.ListCount - 1, 2) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("DocPath").Index).Value)
            lstAgendas.List(lstAgendas.ListCount - 1, 3) = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("PdfPath").Index).Value)
        End If
    Next i
End Sub

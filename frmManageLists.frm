VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmManageLists 
   Caption         =   "UserForm1"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12165
   OleObjectBlob   =   "frmManageLists.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmManageLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    cmdAddEvent.caption = "Add"
    cmdRemoveEvent.caption = "Remove"
    cmdAddCharity.caption = "Add"
    cmdRemoveCharity.caption = "Remove"
    RefreshLists
End Sub

Private Sub cmdAddEvent_Click()
    Dim v As String: v = Trim$(txtNewEvent.value)
    If Len(v) = 0 Then Exit Sub
    AppendToLookup "tblEvents", "Event", v
    txtNewEvent.value = ""
    RefreshLists
End Sub

Private Sub cmdRemoveEvent_Click()
    If lstEvents.ListIndex < 0 Then Exit Sub
    RemoveFromLookup "tblEvents", "Event", CStr(lstEvents.List(lstEvents.ListIndex))
    RefreshLists
End Sub

Private Sub cmdAddCharity_Click()
    Dim v As String: v = Trim$(txtNewCharity.value)
    If Len(v) = 0 Then Exit Sub
    AppendToLookup "tblCharities", "Charity", v
    txtNewCharity.value = ""
    RefreshLists
End Sub

Private Sub cmdRemoveCharity_Click()
    If lstCharities.ListIndex < 0 Then Exit Sub
    RemoveFromLookup "tblCharities", "Charity", CStr(lstCharities.List(lstCharities.ListIndex))
    RefreshLists
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub RefreshLists()
    lstEvents.Clear
    lstCharities.Clear

    Dim loE As ListObject: Set loE = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblEvents")
    If Not loE.DataBodyRange Is Nothing Then
        Dim c As Range
        For Each c In loE.ListColumns(1).DataBodyRange.Cells
            lstEvents.AddItem CStr(c.value)
        Next c
    End If

    Dim loC As ListObject: Set loC = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblCharities")
    If Not loC.DataBodyRange Is Nothing Then
        Dim d As Range
        For Each d In loC.ListColumns(1).DataBodyRange.Cells
            lstCharities.AddItem CStr(d.value)
        Next d
    End If
End Sub

Private Sub AppendToLookup(ByVal tableName As String, ByVal colName As String, ByVal value As String)
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects(tableName)
    Dim lr As ListRow: Set lr = lo.ListRows.Add
    lr.Range.Cells(1, lo.ListColumns(colName).Index).value = value
End Sub

Private Sub RemoveFromLookup(ByVal tableName As String, ByVal colName As String, ByVal value As String)
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects(tableName)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim i As Long
    For i = lo.ListRows.count To 1 Step -1
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns(colName).Index).value) = value Then
            lo.ListRows(i).Delete
            Exit Sub
        End If
    Next i
End Sub


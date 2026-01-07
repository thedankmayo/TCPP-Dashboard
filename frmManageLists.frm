VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmManageLists 
   Caption         =   "UserForm1"
   ClientHeight    =   8120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10035
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
    RefreshLists
End Sub

Private Sub cmdAddEvent_Click()
    On Error GoTo EH
    Dim v As String: v = Trim$(txtNewEvent.value)
    If Len(v) = 0 Then Exit Sub
    AppendToLookup "tblEvents", "Event", v
    txtNewEvent.value = ""
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.AddEvent", Err, v
End Sub

Private Sub cmdRemoveEvent_Click()
    On Error GoTo EH
    If lstEvents.ListIndex < 0 Then Exit Sub
    RemoveFromLookup "tblEvents", "Event", CStr(lstEvents.List(lstEvents.ListIndex))
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.RemoveEvent", Err, ""
End Sub

Private Sub cmdAddCharity_Click()
    On Error GoTo EH
    Dim v As String: v = Trim$(txtNewCharity.value)
    If Len(v) = 0 Then Exit Sub
    AppendToLookup "tblCharities", "Charity", v
    txtNewCharity.value = ""
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.AddCharity", Err, v
End Sub

Private Sub cmdRemoveCharity_Click()
    On Error GoTo EH
    If lstCharities.ListIndex < 0 Then Exit Sub
    RemoveFromLookup "tblCharities", "Charity", CStr(lstCharities.List(lstCharities.ListIndex))
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.RemoveCharity", Err, ""
End Sub

Private Sub cmdAddCOA_Click()
    On Error GoTo EH
    Dim v As String: v = Trim$(txtNewCOA.value)
    If Len(v) = 0 Then Exit Sub
    AppendToLookup "tblCOA", "Category", v
    txtNewCOA.value = ""
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.AddCOA", Err, v
End Sub

Private Sub cmdRemoveCOA_Click()
    On Error GoTo EH
    If lstCOA.ListIndex < 0 Then Exit Sub
    RemoveFromLookup "tblCOA", "Category", CStr(lstCOA.List(lstCOA.ListIndex))
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.RemoveCOA", Err, ""
End Sub

Private Sub cmdAddPaymentMethod_Click()
    On Error GoTo EH
    Dim v As String: v = Trim$(txtNewPaymentMethod.value)
    If Len(v) = 0 Then Exit Sub
    AppendToLookup "tblPaymentMethods", "PaymentMethod", v
    txtNewPaymentMethod.value = ""
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.AddPaymentMethod", Err, v
End Sub

Private Sub cmdRemovePaymentMethod_Click()
    On Error GoTo EH
    If lstPaymentMethods.ListIndex < 0 Then Exit Sub
    RemoveFromLookup "tblPaymentMethods", "PaymentMethod", CStr(lstPaymentMethods.List(lstPaymentMethods.ListIndex))
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.RemovePaymentMethod", Err, ""
End Sub

Private Sub cmdAddTxnType_Click()
    On Error GoTo EH
    Dim v As String: v = Trim$(txtNewTxnType.value)
    If Len(v) = 0 Then Exit Sub
    AppendToLookup "tblTxnTypes", "TxnType", v
    txtNewTxnType.value = ""
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.AddTxnType", Err, v
End Sub

Private Sub cmdRemoveTxnType_Click()
    On Error GoTo EH
    If lstTxnTypes.ListIndex < 0 Then Exit Sub
    RemoveFromLookup "tblTxnTypes", "TxnType", CStr(lstTxnTypes.List(lstTxnTypes.ListIndex))
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.RemoveTxnType", Err, ""
End Sub

Private Sub cmdAddTxnSubtype_Click()
    On Error GoTo EH
    Dim v As String: v = Trim$(txtNewTxnSubtype.value)
    If Len(v) = 0 Then Exit Sub
    AppendToLookup "tblTxnSubtypes", "TxnSubtype", v
    txtNewTxnSubtype.value = ""
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.AddTxnSubtype", Err, v
End Sub

Private Sub cmdRemoveTxnSubtype_Click()
    On Error GoTo EH
    If lstTxnSubtypes.ListIndex < 0 Then Exit Sub
    RemoveFromLookup "tblTxnSubtypes", "TxnSubtype", CStr(lstTxnSubtypes.List(lstTxnSubtypes.ListIndex))
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.RemoveTxnSubtype", Err, ""
End Sub

Private Sub cmdAddMembershipType_Click()
    On Error GoTo EH
    Dim v As String: v = Trim$(txtNewMembershipType.value)
    If Len(v) = 0 Then Exit Sub
    AppendToLookup "tblMembershipTypes", "MembershipType", v
    txtNewMembershipType.value = ""
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.AddMembershipType", Err, v
End Sub

Private Sub cmdRemoveMembershipType_Click()
    On Error GoTo EH
    If lstMembershipTypes.ListIndex < 0 Then Exit Sub
    RemoveFromLookup "tblMembershipTypes", "MembershipType", CStr(lstMembershipTypes.List(lstMembershipTypes.ListIndex))
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.RemoveMembershipType", Err, ""
End Sub

Private Sub cmdAddBoard_Click()
    On Error GoTo EH
    If Len(Trim$(txtBoardName.value)) = 0 Then Exit Sub
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblBoardRoster")
    Dim lr As ListRow: Set lr = lo.ListRows.Add
    lr.Range.Cells(1, lo.ListColumns("Name").Index).value = Trim$(txtBoardName.value)
    lr.Range.Cells(1, lo.ListColumns("Role").Index).value = Trim$(txtBoardRole.value)
    lr.Range.Cells(1, lo.ListColumns("ActiveFlag").Index).value = True
    txtBoardName.value = ""
    txtBoardRole.value = ""
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.AddBoard", Err, ""
End Sub

Private Sub cmdRemoveBoard_Click()
    On Error GoTo EH
    If lstBoardRoster.ListIndex < 0 Then Exit Sub
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblBoardRoster")
    Dim i As Long
    For i = lo.ListRows.count To 1 Step -1
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Name").Index).value) = CStr(lstBoardRoster.List(lstBoardRoster.ListIndex)) Then
            lo.ListRows(i).Delete
            Exit For
        End If
    Next i
    RefreshLists
    Exit Sub
EH:
    modTCPPv2.HandleError "frmManageLists.RemoveBoard", Err, ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub RefreshLists()
    LoadList lstEvents, "tblEvents", "Event"
    LoadList lstCharities, "tblCharities", "Charity"
    LoadList lstCOA, "tblCOA", "Category"
    LoadList lstPaymentMethods, "tblPaymentMethods", "PaymentMethod"
    LoadList lstTxnTypes, "tblTxnTypes", "TxnType"
    LoadList lstTxnSubtypes, "tblTxnSubtypes", "TxnSubtype"
    LoadList lstMembershipTypes, "tblMembershipTypes", "MembershipType"

    lstBoardRoster.Clear
    Dim loB As ListObject: Set loB = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects("tblBoardRoster")
    If Not loB.DataBodyRange Is Nothing Then
        Dim r As Range
        For Each r In loB.ListColumns("Name").DataBodyRange.Cells
            lstBoardRoster.AddItem CStr(r.value)
        Next r
    End If
End Sub

Private Sub LoadList(ByVal lst As MSForms.ListBox, ByVal tableName As String, ByVal colName As String)
    lst.Clear
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("DATA_Lookups").ListObjects(tableName)
    If Not lo.DataBodyRange Is Nothing Then
        Dim c As Range
        For Each c In lo.ListColumns(colName).DataBodyRange.Cells
            lst.AddItem CStr(c.value)
        Next c
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

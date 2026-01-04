VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmManageLists 
   Caption         =   "UserForm1"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12165
   OleObjectBlob   =   "frmManageLists.frx":0000
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstEvents
      Height          =   1200
      Left            =   120
      Top             =   600
      Width           =   2500
   End
   Begin VB.TextBox txtNewEvent
      Height          =   285
      Left            =   120
      Top             =   300
      Width           =   1600
   End
   Begin VB.CommandButton cmdAddEvent
      Caption         =   "Add"
      Height          =   285
      Left            =   1800
      Top             =   300
      Width           =   720
   End
   Begin VB.CommandButton cmdRemoveEvent
      Caption         =   "Remove"
      Height          =   285
      Left            =   1800
      Top             =   1500
      Width           =   720
   End

   Begin VB.ListBox lstCharities
      Height          =   1200
      Left            =   3000
      Top             =   600
      Width           =   2500
   End
   Begin VB.TextBox txtNewCharity
      Height          =   285
      Left            =   3000
      Top             =   300
      Width           =   1600
   End
   Begin VB.CommandButton cmdAddCharity
      Caption         =   "Add"
      Height          =   285
      Left            =   4680
      Top             =   300
      Width           =   720
   End
   Begin VB.CommandButton cmdRemoveCharity
      Caption         =   "Remove"
      Height          =   285
      Left            =   4680
      Top             =   1500
      Width           =   720
   End

   Begin VB.ListBox lstCOA
      Height          =   1200
      Left            =   5880
      Top             =   600
      Width           =   2500
   End
   Begin VB.TextBox txtNewCOA
      Height          =   285
      Left            =   5880
      Top             =   300
      Width           =   1600
   End
   Begin VB.CommandButton cmdAddCOA
      Caption         =   "Add"
      Height          =   285
      Left            =   7560
      Top             =   300
      Width           =   720
   End
   Begin VB.CommandButton cmdRemoveCOA
      Caption         =   "Remove"
      Height          =   285
      Left            =   7560
      Top             =   1500
      Width           =   720
   End

   Begin VB.ListBox lstPaymentMethods
      Height          =   1200
      Left            =   120
      Top             =   2700
      Width           =   2500
   End
   Begin VB.TextBox txtNewPaymentMethod
      Height          =   285
      Left            =   120
      Top             =   2400
      Width           =   1600
   End
   Begin VB.CommandButton cmdAddPaymentMethod
      Caption         =   "Add"
      Height          =   285
      Left            =   1800
      Top             =   2400
      Width           =   720
   End
   Begin VB.CommandButton cmdRemovePaymentMethod
      Caption         =   "Remove"
      Height          =   285
      Left            =   1800
      Top             =   3600
      Width           =   720
   End

   Begin VB.ListBox lstTxnTypes
      Height          =   1200
      Left            =   3000
      Top             =   2700
      Width           =   2500
   End
   Begin VB.TextBox txtNewTxnType
      Height          =   285
      Left            =   3000
      Top             =   2400
      Width           =   1600
   End
   Begin VB.CommandButton cmdAddTxnType
      Caption         =   "Add"
      Height          =   285
      Left            =   4680
      Top             =   2400
      Width           =   720
   End
   Begin VB.CommandButton cmdRemoveTxnType
      Caption         =   "Remove"
      Height          =   285
      Left            =   4680
      Top             =   3600
      Width           =   720
   End

   Begin VB.ListBox lstTxnSubtypes
      Height          =   1200
      Left            =   5880
      Top             =   2700
      Width           =   2500
   End
   Begin VB.TextBox txtNewTxnSubtype
      Height          =   285
      Left            =   5880
      Top             =   2400
      Width           =   1600
   End
   Begin VB.CommandButton cmdAddTxnSubtype
      Caption         =   "Add"
      Height          =   285
      Left            =   7560
      Top             =   2400
      Width           =   720
   End
   Begin VB.CommandButton cmdRemoveTxnSubtype
      Caption         =   "Remove"
      Height          =   285
      Left            =   7560
      Top             =   3600
      Width           =   720
   End

   Begin VB.ListBox lstMembershipTypes
      Height          =   1200
      Left            =   120
      Top             =   4800
      Width           =   2500
   End
   Begin VB.TextBox txtNewMembershipType
      Height          =   285
      Left            =   120
      Top             =   4500
      Width           =   1600
   End
   Begin VB.CommandButton cmdAddMembershipType
      Caption         =   "Add"
      Height          =   285
      Left            =   1800
      Top             =   4500
      Width           =   720
   End
   Begin VB.CommandButton cmdRemoveMembershipType
      Caption         =   "Remove"
      Height          =   285
      Left            =   1800
      Top             =   5700
      Width           =   720
   End

   Begin VB.ListBox lstBoardRoster
      Height          =   1200
      Left            =   3000
      Top             =   4800
      Width           =   2500
   End
   Begin VB.TextBox txtBoardName
      Height          =   285
      Left            =   3000
      Top             =   4500
      Width           =   1200
   End
   Begin VB.TextBox txtBoardRole
      Height          =   285
      Left            =   4260
      Top             =   4500
      Width           =   1200
   End
   Begin VB.CommandButton cmdAddBoard
      Caption         =   "Add"
      Height          =   285
      Left            =   5520
      Top             =   4500
      Width           =   720
   End
   Begin VB.CommandButton cmdRemoveBoard
      Caption         =   "Remove"
      Height          =   285
      Left            =   5520
      Top             =   5700
      Width           =   720
   End

   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   360
      Left            =   9960
      Top             =   9000
      Width           =   960
   End
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
    lr.Range.Cells(1, lo.ListColumns("Name").Index).Value = Trim$(txtBoardName.value)
    lr.Range.Cells(1, lo.ListColumns("Role").Index).Value = Trim$(txtBoardRole.value)
    lr.Range.Cells(1, lo.ListColumns("ActiveFlag").Index).Value = True
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
    For i = lo.ListRows.Count To 1 Step -1
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("Name").Index).Value) = CStr(lstBoardRoster.List(lstBoardRoster.ListIndex)) Then
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
    For i = lo.ListRows.Count To 1 Step -1
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns(colName).Index).value) = value Then
            lo.ListRows(i).Delete
            Exit Sub
        End If
    Next i
End Sub

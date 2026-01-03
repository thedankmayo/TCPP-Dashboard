VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBudget 
   Caption         =   "UserForm1"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5955
   OleObjectBlob   =   "frmBudget.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMonthKey As String

Public Sub InitForMonth(ByVal monthKey As String)
    mMonthKey = monthKey
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo EH
    LoadMonthList
    If Len(mMonthKey) = 0 Then mMonthKey = Format(Date, "yyyy-mm")
    cboMonth.value = mMonthKey
    LoadBudgets
    Exit Sub
EH:
    modTCPPv2.HandleError "frmBudget.Initialize", Err, ""
End Sub

Private Sub cmdLoad_Click()
    mMonthKey = cboMonth.value
    LoadBudgets
End Sub

Private Sub cmdSave_Click()
    On Error GoTo EH
    mMonthKey = cboMonth.value

    modTCPPv2.SetBudget mMonthKey, "Administrative", CDbl(Val(txtBudgetAdministrative.value))
    modTCPPv2.SetBudget mMonthKey, "Programs", CDbl(Val(txtBudgetPrograms.value))
    modTCPPv2.SetBudget mMonthKey, "Fundraising", CDbl(Val(txtBudgetFundraising.value))
    modTCPPv2.SetBudget mMonthKey, "Marketing", CDbl(Val(txtBudgetMarketing.value))
    modTCPPv2.SetBudget mMonthKey, "Travel", CDbl(Val(txtBudgetTravel.value))
    modTCPPv2.SetBudget mMonthKey, "Services", CDbl(Val(txtBudgetServices.value))
    modTCPPv2.SetBudget mMonthKey, "Misc", CDbl(Val(txtBudgetMisc.value))

    modTCPPv2.AuditLog "BudgetSave", "", mMonthKey
    Unload Me
    Exit Sub
EH:
    modTCPPv2.HandleError "frmBudget.Save", Err, mMonthKey
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub LoadBudgets()
    txtBudgetAdministrative.value = Format(modTCPPv2.GetBudget(mMonthKey, "Administrative"), "0.00")
    txtBudgetPrograms.value = Format(modTCPPv2.GetBudget(mMonthKey, "Programs"), "0.00")
    txtBudgetFundraising.value = Format(modTCPPv2.GetBudget(mMonthKey, "Fundraising"), "0.00")
    txtBudgetMarketing.value = Format(modTCPPv2.GetBudget(mMonthKey, "Marketing"), "0.00")
    txtBudgetTravel.value = Format(modTCPPv2.GetBudget(mMonthKey, "Travel"), "0.00")
    txtBudgetServices.value = Format(modTCPPv2.GetBudget(mMonthKey, "Services"), "0.00")
    txtBudgetMisc.value = Format(modTCPPv2.GetBudget(mMonthKey, "Misc"), "0.00")
End Sub

Private Sub LoadMonthList()
    cboMonth.Clear
    Dim i As Long, d As Date
    d = DateSerial(Year(Date), Month(Date), 1)
    For i = -12 To 12
        cboMonth.AddItem Format(DateAdd("m", i, d), "yyyy-mm")
    Next i
End Sub

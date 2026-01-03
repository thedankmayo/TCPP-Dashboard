VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReconcile 
   Caption         =   "UserForm1"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5955
   OleObjectBlob   =   "frmReconcile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReconcile"
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
    If Len(mMonthKey) = 0 Then mMonthKey = Format(Date, "yyyy-mm")
    lblMonth.caption = "Month: " & mMonthKey

    txtBeginningBalance.value = "0"
    txtEndingBalance.value = "0"

    RefreshLedgerTotals
    ComputeDiff
    Exit Sub
EH:
    modTCPPv2.HandleError "frmReconcile.Initialize", Err, mMonthKey
End Sub

Private Sub cmdCompute_Click()
    RefreshLedgerTotals
    ComputeDiff
End Sub

Private Sub cmdSave_Click()
    On Error GoTo EH
    modTCPPv2.SaveReconciliation mMonthKey, CDbl(Val(txtBeginningBalance.value)), CDbl(Val(txtEndingBalance.value))
    Unload Me
    Exit Sub
EH:
    modTCPPv2.HandleError "frmReconcile.Save", Err, mMonthKey
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub RefreshLedgerTotals()
    Dim dep As Double, wd As Double
    modTCPPv2.ComputeMonthLedgerTotals mMonthKey, dep, wd
    lblLedgerDeposits.caption = "Ledger deposits: $" & Format(dep, "0.00")
    lblLedgerWithdrawals.caption = "Ledger withdrawals: $" & Format(wd, "0.00")
End Sub

Private Sub ComputeDiff()
    Dim dep As Double, wd As Double
    modTCPPv2.ComputeMonthLedgerTotals mMonthKey, dep, wd

    Dim b As Double: b = CDbl(Val(txtBeginningBalance.value))
    Dim e As Double: e = CDbl(Val(txtEndingBalance.value))
    Dim expected As Double: expected = b + dep - wd
    Dim diff As Double: diff = expected - e

    lblExpectedEnding.caption = "Expected ending: $" & Format(expected, "0.00")
    lblDifference.caption = "Difference: $" & Format(diff, "0.00")
End Sub

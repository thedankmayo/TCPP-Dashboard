VERSION 5.00
Begin VB.UserForm frmCloseMonth
   Caption         =   "Close Month"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboMonth
      Height          =   285
      Left            =   120
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox txtGateSummary
      Height          =   2200
      Left            =   120
      Top             =   540
      Width           =   5760
      MultiLine       =   -1  'True
   End
   Begin VB.CommandButton cmdCheck
      Caption         =   "Check Gates"
      Height          =   360
      Left            =   1440
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdCloseMonth
      Caption         =   "Close Month"
      Height          =   360
      Left            =   2760
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   360
      Left            =   4440
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "frmCloseMonth"
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
    modTCPPv2.ApplyTheme Me
    LoadMonthList
    If Len(mMonthKey) = 0 Then mMonthKey = Format(Date, "yyyy-mm")
    cboMonth.value = mMonthKey
    RefreshGateSummary
    Exit Sub
EH:
    modTCPPv2.HandleError "frmCloseMonth.Initialize", Err, ""
End Sub

Private Sub cmdCheck_Click()
    RefreshGateSummary
End Sub

Private Sub cmdCloseMonth_Click()
    On Error GoTo EH
    modTCPPv2.CloseMonth cboMonth.value
    RefreshGateSummary
    Exit Sub
EH:
    modTCPPv2.HandleError "frmCloseMonth.CloseMonth", Err, cboMonth.value
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub RefreshGateSummary()
    Dim mc As Long, mr As Long, mAmt As Double, reconOk As Boolean, charityImbalance As Boolean, budgetOverrun As Boolean
    txtGateSummary.value = modTCPPv2.GateCheckMonth(cboMonth.value, mc, mr, mAmt, reconOk, charityImbalance, budgetOverrun)
End Sub

Private Sub LoadMonthList()
    cboMonth.Clear
    Dim i As Long, d As Date
    d = DateSerial(Year(Date), Month(Date), 1)
    For i = -12 To 12
        cboMonth.AddItem Format(DateAdd("m", i, d), "yyyy-mm")
    Next i
End Sub

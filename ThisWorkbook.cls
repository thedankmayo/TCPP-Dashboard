VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub Workbook_Open()
    On Error Resume Next
    modTCPPv2.InitializeTool False
    modTCPPv2.ShowDashboard
End Sub

Public Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    modTCPPv2.CleanupOnClose
End Sub

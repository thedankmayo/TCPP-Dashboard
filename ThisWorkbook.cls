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
    On Error GoTo EH
    modTCPPv2.InitializeTool False
    modTCPPv2.ShowDashboard
    Exit Sub
EH:
    modTCPPv2.HandleError "Workbook_Open", Err, ""
End Sub

Public Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo EH
    modTCPPv2.CleanupOnClose
    Exit Sub
EH:
    modTCPPv2.HandleError "Workbook_BeforeClose", Err, ""
End Sub

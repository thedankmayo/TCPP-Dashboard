VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImports 
   Caption         =   "UserForm1"
   ClientHeight    =   2060
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8070
   OleObjectBlob   =   "frmImports.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    cboSource.Clear
    cboSource.AddItem "Zeffy"
    cboSource.AddItem "Blaze"
    cboSource.value = "Zeffy"
End Sub

Private Sub cmdBrowse_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = False
    fd.Title = "Select import CSV"
    fd.Filters.Clear
    fd.Filters.Add "CSV", "*.csv"
    If fd.Show <> -1 Then Exit Sub
    txtFilePath.value = fd.SelectedItems(1)
End Sub

Private Sub cmdImport_Click()
    On Error GoTo EH
    If Len(Trim$(txtFilePath.value)) = 0 Then Exit Sub
    modTCPPv2.ImportCsvRaw cboSource.value, txtFilePath.value
    MsgBox "Import staged.", vbInformation
    Exit Sub
EH:
    modTCPPv2.HandleError "frmImports.Import", Err, txtFilePath.value
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

VERSION 5.00
Begin VB.UserForm frmImports
   Caption         =   "Imports"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboSource
      Height          =   285
      Left            =   120
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox txtFilePath
      Height          =   285
      Left            =   1440
      Top             =   120
      Width           =   3600
   End
   Begin VB.CommandButton cmdBrowse
      Caption         =   "Browse"
      Height          =   285
      Left            =   5160
      Top             =   120
      Width           =   720
   End
   Begin VB.CommandButton cmdImport
      Caption         =   "Import"
      Height          =   360
      Left            =   120
      Top             =   600
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   360
      Left            =   1440
      Top             =   600
      Width           =   1200
   End
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

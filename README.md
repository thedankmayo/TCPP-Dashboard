# TCPP-Dashboard
A excel-based dashboard to be used for managing all things treasury related for small local no-profits.

## Quick start (one-click ready)
The tool is designed to be one-click ready **after import** as long as:
- Macros are enabled in Excel.
- The Word templates are stored next to the workbook.

On first open, the workbook runs `Workbook_Open` → `InitializeTool(False)` which creates all sheets, tables, and default folders automatically.

## Required files (same folder as the .xlsm)
- `TCPP Board Meeting Minutes Template.docx`
- `Template Meeting Agenda.docx`

## Importing the VBA code into Excel
### Option A: Manual import
1. Open Excel and your `.xlsm` workbook.
2. Press **Alt+F11** to open the VBA editor.
3. **File → Import File…** and import all `.bas`, `.frm`, and `.cls` files in this repo.
4. Save the workbook as **Macro-Enabled (.xlsm)** and reopen it.

### Option B: One-command PowerShell import
> Requires Excel installed and **Trust access to the VBA project object model** enabled.

```powershell
./import-vba.ps1 -WorkbookPath "C:\Path\To\TCPP_Dashboard.xlsm"
```

This script will import all modules and forms, overwrite the `ThisWorkbook` and `Sheet#` code, and save the workbook.

## First run checklist
1. Open the workbook and **enable macros**.
2. The dashboard should appear automatically.
3. Use **RunSelfTest** on the dashboard to confirm tables and paths.

## Default folders (auto-created)
```
.\BoardPackets\
.\Minutes\DOCX\
.\Minutes\PDF\
.\Agenda\DOCX\
.\Agenda\PDF\
.\Imports\Zeffy\
.\Imports\Blaze\
.\Reports\
```

## Notes
- If you move the workbook, keep the templates next to it or update the paths in `tblConfig`.
- All data sheets are `xlSheetVeryHidden`; the dashboard is the only user-facing surface.

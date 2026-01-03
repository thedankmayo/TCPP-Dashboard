param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,
    [string]$SourcePath = (Get-Location).Path
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $WorkbookPath)) {
    Write-Host "Workbook not found. Creating new workbook at $WorkbookPath"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Add()
    $xlOpenXMLWorkbookMacroEnabled = 52
    $workbook.SaveAs($WorkbookPath, $xlOpenXMLWorkbookMacroEnabled)
    $workbook.Close($true)
    $excel.Quit()
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($WorkbookPath)
$project = $workbook.VBProject

function Import-DocumentModule {
    param(
        [string]$ComponentName,
        [string]$FilePath
    )
    $component = $project.VBComponents.Item($ComponentName)
    $code = $component.CodeModule
    $code.DeleteLines(1, $code.CountOfLines)
    $code.AddFromFile($FilePath)
}

function Remove-IfExists {
    param([string]$ComponentName)
    try {
        $component = $project.VBComponents.Item($ComponentName)
        if ($component.Type -ne 100) {
            $project.VBComponents.Remove($component)
        }
    } catch {
        # no-op
    }
}

Get-ChildItem -Path $SourcePath -Filter "*.bas" | ForEach-Object {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
    Remove-IfExists $name
    $project.VBComponents.Import($_.FullName) | Out-Null
}

Get-ChildItem -Path $SourcePath -Filter "*.frm" | ForEach-Object {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
    Remove-IfExists $name
    $project.VBComponents.Import($_.FullName) | Out-Null
}

Get-ChildItem -Path $SourcePath -Filter "*.cls" | ForEach-Object {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
    if ($name -eq "ThisWorkbook") {
        Import-DocumentModule -ComponentName "ThisWorkbook" -FilePath $_.FullName
    } elseif ($name -match "^Sheet\d+$") {
        Import-DocumentModule -ComponentName $name -FilePath $_.FullName
    } else {
        Remove-IfExists $name
        $project.VBComponents.Import($_.FullName) | Out-Null
    }
}

$workbook.Save()
$workbook.Close($true)
$excel.Quit()

Write-Host "Import completed. Open the workbook and enable macros to initialize the tool."

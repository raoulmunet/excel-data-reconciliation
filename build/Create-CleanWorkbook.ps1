#requires -Version 5.1
<#
.SYNOPSIS
Builds a CLEAN macro-enabled workbook from the three production VBA modules.
.DESCRIPTION
Requires Windows Excel Desktop (2019 reference) and temporary authorization to
"Trust access to the VBA project object model" in Excel Trust Center.
Does not change macro security settings automatically.
Open source only; do not run scripts from untrusted repositories.
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'ExcelDataReconciliation.xlsm')
)
$ErrorActionPreference = 'Stop'
$modulePaths = @(
    (Join-Path $PSScriptRoot 'src\modAdvancedReconciliation.bas'),
    (Join-Path $PSScriptRoot 'src\modReportPresentation.bas'),
    (Join-Path $PSScriptRoot 'src\modUnifiedDashboard.bas')
)
foreach ($path in $modulePaths) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Missing source module: $path. Download/clone the complete repository."
    }
}
if (Test-Path -LiteralPath $OutputPath) {
    throw "Output already exists: $OutputPath. Rename or back up the old file first. Nothing was overwritten."
}

$excel = $null
$book = $null
$successful = $false
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    try {
        $project = $book.VBProject
        $components = $project.VBComponents
        foreach ($path in $modulePaths) {
            $component = $components.Import($path)
            Write-Host ("Imported: " + $component.Name)
        }
    } catch {
        throw "VBA module import failed. In Excel: File > Options > Trust Center > Trust Center Settings > Macro Settings > enable 'Trust access to the VBA project object model' temporarily, then restart Excel. Details: $($_.Exception.Message)"
    }
    # 52 = xlOpenXMLWorkbookMacroEnabled
    $book.SaveAs($OutputPath, 52)
    $excel.Run("'" + $book.Name.Replace("'", "''") + "'!BuildUnifiedDashboard")
    # Remove only the blank default worksheet(s) from the newly-created workbook.
    for ($i = $book.Worksheets.Count; $i -ge 1; $i--) {
        $sheet = $book.Worksheets.Item($i)
        if ($sheet.Name -ne 'Dashboard') {
            $sheet.Delete()
        }
    }
    $book.Save()
    $successful = $true
    Write-Host "Created: $OutputPath"
    Write-Host "Open it in Excel, compile via Debug > Compile VBAProject, and run the sample tests."
}
finally {
    if ($null -ne $book) {
        try { $book.Close($successful) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) } catch { }
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

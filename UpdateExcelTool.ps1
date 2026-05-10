param (
    [Parameter(Mandatory=$true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory=$true)]
    [string]$SheetName,

    [Parameter(Mandatory=$true)]
    [string]$CellAddress,

    [Parameter(Mandatory=$true)]
    [string]$Formula
)

$excel = $null
$workbook = $null
$sheet = $null

try {
    # Resolve full path
    $fullPath = (Resolve-Path $WorkbookPath).Path

    # Create Excel COM object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false

    # ----------------------------
    # Open Workbook
    # ----------------------------
    try {
        $workbook = $excel.Workbooks.Open($fullPath, 0, $false)

        if ($workbook.ReadOnly) {
            throw "Workbook opened as ReadOnly"
        }
    }
    catch {
        throw "Failed to open workbook '$fullPath'. $($_.Exception.Message)"
    }

    # ----------------------------
    # Get Worksheet
    # ----------------------------
    try {
        $sheet = $workbook.Worksheets.Item($SheetName)
    }
    catch {
        throw "Failed to access worksheet '$SheetName'. $($_.Exception.Message)"
    }

    # ----------------------------
    # Apply Formula
    # ----------------------------
    try {
        $sheet.Range($CellAddress).Formula = $Formula
    }
    catch {
        throw "Failed to write formula to cell '$CellAddress'. $($_.Exception.Message)"
    }

    # ----------------------------
    # Save Workbook
    # ----------------------------
    try {
        $workbook.Save()
    }
    catch {
        throw "Failed to save workbook. $($_.Exception.Message)"
    }

    Write-Host "Successfully updated $CellAddress in $SheetName in $fullPath" -ForegroundColor Green
}
catch {
    Write-Error $_.Exception.Message
}
finally {
    if ($workbook) {
        $workbook.Close($false)
    }

    if ($excel) {
        $excel.Quit()
    }

    # Release COM objects
    if ($sheet) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
    }

    if ($workbook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }

    if ($excel) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}


#testcase with cmd
#powershell -ExecutionPolicy Bypass -File "C:\t\UpdateExcelTool.ps1" -WorkbookPath "C:\t\1.xlsx" -SheetName "Sheet1" -CellAddress "A1" -Formula "=A2+A3"

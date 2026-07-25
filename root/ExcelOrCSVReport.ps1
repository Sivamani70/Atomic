<#
.SYNOPSIS
    Excel or CSV report generation module for Threat Intelligence IOC reputation data.

.DESCRIPTION
    This script converts dynamic collections of Indicator of Compromise (IOC) reputation data 
    into a cleanly formatted Microsoft Excel (.xlsx) workbook using COM automation.
    If Microsoft Excel is not installed on the system, it gracefully falls back to generating a standard .csv file.

    NOTE: This script is not intended to be executed directly by the user. It acts as 
    a specialized module invoked by wrapper scripts like AbuseIPDB Script and SOLO.ps1.
.NOTES
    Author: SivaMani70
    Date: May 2026

    Dependencies: Relies on an external file located at $PSScriptRoot\Prompts.ps1 which must contain the Get-Prompt function.
    
#>



. $PSScriptRoot\Prompts.ps1

function New-Report([string]$IOCType, [System.Collections.Generic.List[PSCustomObject]]$Data) {
    if ((Test-Path -Path HKLM:\SOFTWARE\Microsoft\Office\*\Excel\)) {
        Write-Host "Excel Application found"
        Write-Host "Creating Excel File"
        New-XL -IOCType $IOCType -SheetName "$IOCType Reputation" -Data $Data
        return
    }
    Write-Host "Excel Application not found"
    Write-Host "Creating CSV File"
    New-CSV -IOCType $IOCType -Data $Data
}

function Get-FileName([string]$IOCType) {
    $ChosenPrompt = Get-Prompt
    Write-Host -Object $ChosenPrompt -ForegroundColor Yellow
    [String]$FileName = Read-Host -Prompt "`n> Enter File Name`t" 
    if ([string]::IsNullOrWhiteSpace($FileName)) { $FileName = "Untitled_Chaos" }
    return "$IOCType Rep - $FileName"
}    

function New-CSV([string]$IOCType, [System.Collections.Generic.List[PSCustomObject]]$Data) {
    [String]$FileName = Get-FileName -IOCType $IOCType
    Write-Host "Creating $FileName.csv file"
    $Data | Export-Csv -Path "$FileName.csv" -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "Completed creating $FileName.csv" 
}

function New-XL([string]$IOCType, [string]$SheetName, [System.Collections.Generic.List[PSCustomObject]]$Data) {
    [String]$FileName = Get-FileName -IOCType $IOCType
    $FullName = (Get-Location).Path + "\$FileName.xlsx"
    Write-Host "Creating $FullName file"
    $Excel = New-Object -ComObject Excel.Application
    $WorkBook = $Excel.Workbooks.Add()
    try {
        $RepSheet = $WorkBook.Sheets["Sheet1"]
        $RepSheet.Name = $SheetName

        $Properties = $Data[0].PSObject.Properties.Name
        $Row = 1
        $Col = 1
            
        foreach ($Property in $Properties) {
            $HeaderCell = $RepSheet.Cells[$Row, $Col]
            Format-Header -Cell $HeaderCell -Text $Property
            Set-CenterContent -Cell $HeaderCell
            $Col += 1
        }
            
        $Row = 2
        foreach ($Obj in $Data) {
            $Col = 1
            foreach ($Property in $Properties) {
                $Cell = $RepSheet.Cells[$Row, $Col]
                Write-Cell -Cell $Cell -Value "$($Obj.$Property)"
                $Col += 1
            }
            $Row += 1
        }
        $RepSheet.Columns.AutoFit()
        $WorkBook.SaveAs($FullName)
        if (Test-Path -Path $FullName) {
            Write-Host "File created successfully: $FullName" -ForegroundColor Green
        }
        else {
            Write-Warning "File creation failed: $FullName"
        }
    }
    finally {
        if ($null -ne $WorkBook) {
            # Set DisplayAlerts to false so it doesn't pop up a "Save changes?" window if it crashes
            $Excel.DisplayAlerts = $false
            $WorkBook.Close($false) # False means don't save pending changes on a crash
            $Excel.DisplayAlerts = $true
        }
        if ($null -ne $Excel) {
            $Excel.Quit()
            $ExitCode = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
            $Excel = $null
            Write-Host "Exit-Code: $ExitCode" -ForegroundColor Yellow
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}

function Format-Header($Cell, [string]$Text) {
    Write-Cell -Cell $Cell -Value $Text
    $Cell.Interior.ColorIndex = 37 # Light Blue
    $Cell.Font.Bold = $true        # Professional touch
}
    
function Set-CenterContent($Cell) {
    $Cell.HorizontalAlignment = 3 # Center
}

function Write-Cell($Cell, $Value) {
    $Cell.Value = $Value
    $Cell.Borders.LineStyle = 1
    $Cell.Borders.ColorIndex = 1
}
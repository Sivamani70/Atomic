. $PSScriptRoot\Prompts.ps1

function Get-FileName() {
    $ChosenPrompt = Get-Prompt
    Write-Host -Object $ChosenPrompt -ForegroundColor Yellow
    [String]$FileName = Read-Host -Prompt "`n> Enter File Name`t" 
    if ([string]::IsNullOrWhiteSpace($FileName)) { $FileName = "Untitled_Chaos" }
    return "Reputation Report - $FileName"
}

function New-ExcelReport([System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[PSCustomObject]]]$Data) { 
    [String]$FileName = Get-FileName
    [String]$FullName = (Get-Location).Path + "\$FileName.xlsx"
    Write-Host "Creating $FullName file"
    $Excel = New-Object -ComObject Excel.Application
    $WorkBook = $Excel.Workbooks.Add()

    try {

        foreach ($IOCType in $Data.keys) {
            $WorkSheet = $WorkBook.Worksheets.Add()
            $SheetName = "$IOCType Reputation"
            $WorkSheet.Name = $SheetName

            Write-Host "Creating Reputation sheet for $IOCType data..." -ForegroundColor Green
            $IOCData = $Data[$IOCType]
            $Properties = $IOCData[0].PSObject.Properties.Name
            $Row = 1
            $Col = 1

            foreach ($Property in $Properties) {
                $HeaderCell = $WorkSheet.Cells[$Row, $Col]
                Format-Header -Cell $HeaderCell -Text $Property
                Set-CenterContent -Cell $HeaderCell
                $Col += 1
            }

            $Row = 2
            foreach ($Obj in $IOCData) {
                $Col = 1
                foreach ($Property in $Properties) {
                    $Cell = $WorkSheet.Cells[$Row, $Col]
                    Write-Cell -Cell $Cell -Value "$($Obj.$Property)"
                    $Col += 1
                }
                $Row += 1
            }
            $WorkSheet.Columns.AutoFit() | Out-Null
        }
        $WorkBook.Sheets["Sheet1"].Delete()
        $WorkBook.SaveAs($FullName)

        if (Test-Path -Path $FullName) {
            Write-Host "File created successfully: $FullName" -ForegroundColor Green
        }
        else {
            Write-Warning "File creation failed: $FullName"
        }
    }
    catch {
        Write-Error "Error creating Excel Report: $_"
    }
    finally {
        if ($null -ne $WorkBook) {
            $Excel.DisplayAlerts = $false
            $WorkBook.Close($false)
            $Excel.DisplayAlerts = $true
        }
        $ExitCode = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
        $Excel = $null
        Write-Host "Exit-Code: $ExitCode" -ForegroundColor Yellow
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
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
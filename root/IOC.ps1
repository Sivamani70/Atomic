<#
.SYNOPSIS
    Indicator of Compromise (IOC) extraction and sanitization module.

.DESCRIPTION
    This script provides the `IOC` class along with procedural wrapper functions (`Get-IOC` and `Get-IOC_TXT`) to extract, sanitize and deduplicate threat indicators 
    from both Microsoft Excel workbooks (.xlsx) and plain-text (.txt) files.

    Key Features:
    - Automatically de sanitizes indicators (e.g., converting `[.]`, `[:]`, `[://]` back to standard formats).
    - Removes wrapping brackets and parentheses `(`, `)`, `[`, `]`.
    - Filters out standard spreadsheet table headers (e.g., "hashes", "ip address", "url") using a built-in blacklist.
    - Guarantees unique outputs by aggregating results into a `HashSet[String]`.
    - Skips non-IOC sheets like "Techniques and Tactics" automatically.

.OUTPUTS
    [System.Collections.Generic.HashSet[String]] Set of unique, de sanitized IOCs.

.NOTES
    Author: SivaMani70
    Date: May 2026

    Not meant to be executed directly. Use wrapper functions `Get-IOC` or `Get-IOC_TXT` for proper invocation.
    As of now using inside the TMC, IOC_90Day_Aggregator.ps1 and inside virustotal scripts to extract IOCs from Excel and TXT files.
    Dependencies: Requires Microsoft Excel COM Object (`Excel.Application`) for workbook processing.
#>



function Get-IOC {
    param (
        [Parameter(Mandatory = $true)]
        [string]$WBPath,
        [Parameter(Mandatory = $true)]
        [System.Object]$Excel
    )  

    [IOC]$IOCProcessor = [IOC]::new($Excel, $WBPath)
    return $IOCProcessor.ExtractIOCs()
}

function Get-IOC_TXT {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )  
    [IOC]$IOCProcessor = [IOC]::new($FilePath)
    return $IOCProcessor.ExtractIOC_TXT()
}

class IOC {
    [System.Object]$Excel
    [string]$WorkBookPath
    [string]$FilePath
    [System.Collections.Generic.HashSet[String]]$ListOfIOCs
    [string[]]$Blacklist = @('hashes', 'md5', 'Sha1', 'sha256', 'hash', 'domain', 'domains', 'url', 'urls', 'email', 'emails', 'ip', 'ips', 'ip address')

    IOC([string]$Path) {
        $this.FilePath = $Path        
        $this.ListOfIOCs = New-Object System.Collections.Generic.HashSet[String]
    }

    IOC([System.Object]$E, [string]$Path) {
        $this.Excel = $E
        $this.WorkBookPath = $Path
        $this.ListOfIOCs = New-Object System.Collections.Generic.HashSet[String]
    }

    [System.Collections.Generic.HashSet[String]] ExtractIOCs() {
        $WorkBook = $this.Excel.Workbooks.Open($this.WorkBookPath)
        Write-Host "Extracting IOCs from the WorkBook: [$($this.WorkBookPath)]"  -ForegroundColor Green

        try {
            foreach ($Sheet in $WorkBook.Sheets) {

                [string]$SheetName = $Sheet.Name
                if ($SheetName.ToLower() -eq "Techniques and Tactics".ToLower()) { 
                    Write-Warning "Skipping the Sheet -- [$SheetName]"   
                    continue; 
                }

                $WorkSheet = $WorkBook.Sheets[$SheetName]
                Write-Host "Extracting IOCs from the Sheet: [$SheetName]"  -ForegroundColor Green
                $Data = $WorkSheet.UsedRange.Value2
                
                foreach ($Cell in $Data) {
                    # Skip if the cell is null or just empty space
                    if ([string]::IsNullOrWhiteSpace($Cell)) { continue }
                    $Indicator = ([string]$Cell).ToLower().Trim().Replace("[:]", ":").Replace("[://]", "://").Replace("[.]", ".")
                    if ($Indicator -in $this.Blacklist) {
                        continue
                    }
                    Write-Host "Cell value: $Indicator" -BackgroundColor Black -ForegroundColor DarkCyan
                    $this.ListOfIOCs.Add($Indicator) | Out-Null
                }
            }
        }
        finally {
            $Workbook.Close()
        }

        return $this.ListOfIOCs
    }

    [System.Collections.Generic.HashSet[String]] ExtractIOC_TXT() {
        $Content = Get-content -path $this.FilePath
        foreach ($Indicator in $Content) {
            if ([string]::IsNullOrWhiteSpace($Indicator)) { continue }
            $Indicator = ([string]$Indicator).ToLower().Trim().Replace("[:]", ":").Replace("[://]", "://").Replace("[.]", ".").Replace("(", "").Replace(")", "").Replace("[", "").Replace("]", "")
            if ($Indicator -in $this.Blacklist) {
                continue
            }
            $this.ListOfIOCs.Add($Indicator) | Out-Null
        }
        return $this.ListOfIOCs
    }   
}
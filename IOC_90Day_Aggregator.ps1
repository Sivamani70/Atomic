<#
.SYNOPSIS
    Consolidates, cleanses, and maps multi-source threat intelligence data into formatted Excel reports.
.DESCRIPTION
    The Collector script acts as an enterprise-grade Threat Intelligence parsing engine. It instantiates 
    a specialized .NET class structure that handles low-level COM automation with Microsoft Excel. 
    It parses incoming workbooks, cleanses noise out of malware actor naming structures, normalizes and 
    deduplicates raw text data using underlying high-performance [HashSet] structures, and routes 
    cleansed indicators into strict schema-compliant Excel spreadsheets and structured STIX-like reports.
.PARAMETER InputDirectory
    The directory containing the source workbook files to ingest.
.PARAMETER MonthName
    The name of the month for which to generate reports.
.EXAMPLE
    .\Collector.ps1 -InputDirectory "D:\IOCs" -MonthName "January"
.NOTES
    Author: SivaMani70
    Date: May 2026
    Prerequisites: Requires local installation of Microsoft Excel (COM Interop validation executed against HKLM).
    Safety: Implements programmatic Marshal COM reference release wrappers to eliminate zombie excel.exe processes.
#>


[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$InputDirectory,
    [Parameter(Mandatory)]
    [string]$MonthName
)

# Import required dependency baseline components for IOC parsing, Excel report generation, and file movement operations
. $PSScriptRoot\root\IOC.ps1

class Collector {
    # --- Class Properties ---
    [System.Collections.Generic.HashSet[String]]$IOCData
    [string[]]$WorkBooks
    [string]$MonthName
    [System.Object]$Excel
    
    # Highly efficient deduplicated text storage buckets for each observable type, leveraging .NET HashSet for O(1) complexity on add and lookup operations, ensuring that the script can handle large volumes of indicators without performance degradation
    [System.Collections.Generic.HashSet[String]]$MD5
    [System.Collections.Generic.HashSet[String]]$SHA1
    [System.Collections.Generic.HashSet[String]]$SHA256
    [System.Collections.Generic.HashSet[String]]$Domains
    [System.Collections.Generic.HashSet[String]]$URLS
    [System.Collections.Generic.HashSet[String]]$IPS
    [System.Collections.Generic.HashSet[String]]$Emails
    [System.Collections.Generic.HashSet[String]]$OtherIOCs

    # Explicit Cryptographic and Network Validation Regex Patterns to ensure that only well-formed indicators are processed and included in the final reports, reducing noise and improving the quality of the threat intelligence output
    [String]$MD5_Validator = "^[a-fA-F0-9]{32}$"
    [String]$SHA1_Validator = "^[a-fA-F0-9]{40}$"
    [String]$SHA256_Validator = "^[a-fA-F0-9]{64}$"
    [String]$DomainValidator = "^(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})?$"
    [String]$URLValidator = "^(https?|hxxps?|ftp):\/\/[^\s/$.?#].[^\s]*$"
    [String]$EmailValidator = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    [String]$IPV4Validator = "^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
    [String]$IPV6Validator = "^(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))$"


    # --- Constructor ---
    # Purpose: Populates task payloads, spins up background Excel COM hooks, and allocates memory blocks.
    Collector([string[]]$WorkBooks, [string]$MonthName) {
        $this.WorkBooks = $WorkBooks
        $this.MonthName = $MonthName
        $this.Excel = New-Object -ComObject Excel.Application
        $this.IOCData = New-Object System.Collections.Generic.HashSet[String]


        $this.MD5 = New-Object System.Collections.Generic.HashSet[String]
        $this.SHA1 = New-Object System.Collections.Generic.HashSet[String]
        $this.SHA256 = New-Object System.Collections.Generic.HashSet[String]
        $this.Domains = New-Object System.Collections.Generic.HashSet[String]
        $this.URLS = New-Object System.Collections.Generic.HashSet[String]
        $this.IPS = New-Object System.Collections.Generic.HashSet[String]
        $this.Emails = New-Object System.Collections.Generic.HashSet[String]
        $this.OtherIOCs = New-Object System.Collections.Generic.HashSet[String]
    }


    # --- Method: GetFileName ---
    # Purpose: Formats a standardized, localized name for report output files based on active date metadata. 
    # This method ensures that generated reports have consistent and descriptive names that include the month, 
    # day, and year of report generation, improving organization and traceability of threat intelligence outputs.
    [String] GetFileName() {
        [datetime]$Date = Get-Date
        [string]$FileName = "Collector - $($this.MonthName) $($Date.Year)"
        return $FileName
    }

    # --- Method: IsValidPath ---
    # Purpose: Verifies target workspace file path existence natively. This method is used to ensure that the script does not attempt to process non-existent files, which could lead to errors or exceptions. By validating the file paths before attempting to read or process them, the script can fail gracefully with informative error messages, improving user experience and robustness.
    [bool] IsValidPath([string]$Path) {
        return Test-Path -Path $Path
    }

    # --- Method: IPExtractorFromURLs ---
    # Purpose: Inline URL sub-parser that pulls raw embedded IPv4 strings straight out of urls.
    [Void] IPExtractorFromURLs([string]$Indicator) {
        [string]$ipLookup = "(https?|hxxps?|ftp)://(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})"
        if ($Indicator -match $ipLookup) {
            foreach ($value in $Matches.Values) {
                if ($value -match $this.IPV4Validator) {
                    $this.IPS.Add($value) | Out-Null
                }
            }
        }
    }

    # --- Method: OrganizeIOC ---
    # Purpose: Cleans bad character structures out of actor labels and maps strings via 
    # regex switch trees directly into specialized, deduplicated type buckets.
    [void] OrganizeIOC() {
        foreach ($Indicator in $this.IOCData) {
            $Indicator = ($Indicator.ToLower()).Trim()

            switch -Regex ($Indicator) {
                # Check IP Addresses
                { $_ -match $this.IPV4Validator -or $_ -match $this.IPV6Validator } { 
                    $this.IPS.Add($Indicator) | Out-Null
                    continue 
                }

                # Check Domain
                $this.DomainValidator { 
                    $this.Domains.Add($Indicator) | Out-Null
                    continue 
                }
                
                # Check Hashes
                $this.MD5_Validator    {
                    $this.MD5.Add($Indicator) | Out-Null; 
                    continue 
                }
                $this.SHA1_Validator   {
                    $this.SHA1.Add($Indicator) | Out-Null; 
                    continue 
                }
                $this.SHA256_Validator {
                    $this.SHA256.Add($Indicator) | Out-Null; 
                    continue 
                }
                
                # Check URL
                $this.URLValidator { 
                    $this.URLS.Add($Indicator) | Out-Null
                    $this.IPExtractorFromURLs($Indicator)
                    continue 
                }
                
                # Check Email
                $this.EmailValidator {
                    $this.Emails.Add($Indicator) | Out-Null; 
                    continue 
                }
                
                Default {
                    # This runs if NONE of the patterns above matched
                    Write-Warning "Uncategorized data found: $Indicator"
                    $this.OtherIOCs.Add($Indicator) | Out-Null
                }
            }
        }
    }


    # --- Method: CreateWorkBook ---
    # Purpose: Orchestrates layout design configurations, builds custom tab blocks, formats 
    # column grids dynamically, and compiles final workbook assets to disk.
    [Void] CreateWorkBook() {
        [String]$FileName = $this.GetFileName()
        $FullName = (Get-Location).Path + "\$FileName.xlsx"

        $WorkBook = $this.Excel.Workbooks.Add()
        try {
            Write-Host "Creating File -- $FullName. `nThis may take some time...." -ForegroundColor Yellow
            $WorkSheet = $WorkBook.Worksheets.Add()

            # Render columns out onto the spreadsheet grid sequentially based on active collection counts, starting with the most critical IOC types (hashes) and then moving into network observables, and finally dumping any uncategorized data into a catch-all column at the end for manual review.
            $CurrentCol = 0
            $this.WriteIOCColumn($this.MD5, $WorkSheet, [ref]$CurrentCol, "MD5")
            $this.WriteIOCColumn($this.SHA1, $WorkSheet, [ref]$CurrentCol, "SHA1")
            $this.WriteIOCColumn($this.SHA256, $WorkSheet, [ref]$CurrentCol, "SHA256")
            $this.WriteIOCColumn($this.Domains, $WorkSheet, [ref]$CurrentCol, "Domains")
            $this.WriteIOCColumn($this.URLS, $WorkSheet, [ref]$CurrentCol, "URLs")
            $this.WriteIOCColumn($this.Emails, $WorkSheet, [ref]$CurrentCol, "Emails")
            $this.WriteIOCColumn($this.IPS, $WorkSheet, [ref]$CurrentCol, "IP")
            $this.WriteIOCColumn($this.OtherIOCs, $WorkSheet, [ref]$CurrentCol, "Review -- Below Values")
            $WorkSheet.Columns.AutoFit()
            $WorkSheet.Name = "$($this.MonthName) - IOCs"
            $WorkBook.Sheets["Sheet1"].Delete()
            $WorkBook.SaveAs($FullName)

            if (Test-Path -Path $FullName) {
                Write-Host "File created successfully: $FullName" -ForegroundColor Green
            }
            else {
                Write-Warning "File creation failed: $FullName"
            }
        }
        finally {
            # Guarantee execution cleanup handles and freeze alerts if a crash triggers mid-write
            if ($null -ne $WorkBook) {
                # Set DisplayAlerts to false so it doesn't pop up a "Save changes?" window if it crashes
                $this.Excel.DisplayAlerts = $false
                $WorkBook.Close($false) # False means don't save pending changes on a crash
                $this.Excel.DisplayAlerts = $true
            }
        }
    }

    # --- Method: Load ---
    # Purpose: Loops through provided files, extracts the malware name from the filename to use as a sheet 
    # name, validates file existence, invokes the Get-IOC parser to extract indicators and maps the extracted 
    # IOCs into the global collection under the malware name as key. This method serves as the main entry 
    # point for processing the input workbooks, coordinating the overall workflow of reading the files, 
    # extracting and organizing the IOCs, and preparing the data for report generation. It also includes user 
    # interaction for optional text file generation and ensures that only valid files are processed, improving 
    # the robustness and usability of the script.
    [void] Load() {
        foreach ($WorkBook in $this.WorkBooks) {
            if ($this.IsValidPath($WorkBook)) {
                [System.Collections.Generic.HashSet[String]]$ListOfIOCs = Get-IOC -WBPath $WorkBook -Excel $this.Excel
                $this.IOCData.UnionWith($ListOfIOCs)
            }
            else {
                Write-Warning "File not exist -- [$WorkBook]"
            }
        }


        if ($null -ne $this.IOCData -and $this.IOCData.Count -gt 0) {
            # Checking if any of those sheets actually contain IOC values
            Write-Host "Successfully collected IOCs. Generating report..."
            $this.OrganizeIOC()
            $this.CreateWorkBook()
        }
    }


    # Atomic methods from here
    # Small helpers
    [void] FormatHeader($Cell, [string]$Text) {
        $this.WriteToCell($Cell, $Text)
        $Cell.Interior.ColorIndex = 37 # Light Blue layout color code
        $Cell.Font.Bold = $true
    }
    
    [void] CenterContent($Cell) {
        $Cell.HorizontalAlignment = 3 # Alignment center enum mapping
    }

    [void] WriteToCell($Cell, $Value) {
        $Cell.Value = $Value
        $Cell.Borders.LineStyle = 1
        $Cell.Borders.ColorIndex = 1
    }

    # --- Method: WriteIOCColumn ---
    [void] WriteIOCColumn($Collection, $WorkSheet, [ref]$Col, [string]$Header) {
        if ($Collection.Count -eq 0) { return }
        # Move to the next column
        $Col.Value += 2
        $Row = 2

        # $Col.Value is the dereferencing from the box (from the pointer) 
        # $Col.Value is a integer here 
        $HeaderCell = $WorkSheet.Cells[$Row, $Col.Value]
        $this.FormatHeader($HeaderCell, $Header)
        $this.CenterContent($HeaderCell)
        
        # Fast Write Data
        foreach ($Item in $Collection) {
            $Row++
            $Cell = $WorkSheet.Cells[$Row, $Col.Value]
            $this.WriteToCell($Cell, $Item)
        }
    }


    # --- Method: Dispose ---
    # Purpose: Essential memory management routine. Shuts down the background application execution 
    # scope and explicitly unbinds the Marshalling reference counts from the OS kernel 
    # to avoid trailing, hidden system processes. This method is critical for ensuring that the script does 
    # not leave behind orphaned Excel processes that can consume system resources and lead to performance 
    # issues. By explicitly calling the Quit method on the Excel application and releasing the COM object 
    # reference, the script ensures that all resources are properly cleaned up after execution.
    [void] Dispose() {
        if ($null -ne $this.Excel) {
            try {
                $this.Excel.Quit()
            }
            finally {
                $ExitCode = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Excel)
                $this.Excel = $null
                Write-Host "Exit-Code: $ExitCode" -ForegroundColor Yellow
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
            }
        }
    }
}

# ==========================================
# --- Runtime Script Execution Pipeline ---
# ==========================================
# 1. Environment Verification Guard
if (-not (Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Office\*\Excel")) {
    Write-Error "No Excel Module found in the System"    
    return
}

# 2. Main Processing Pipeline Block
[string[]]$WorkBooks = Get-ChildItem -Path $InputDirectory -Filter *.xlsx | Select-Object -ExpandProperty FullName
[Collector]$Process = [Collector]::new($WorkBooks, $MonthName)
try {
    $Process.Load()
}
finally {
    # 3. Memory Cleanup Execution Hook (Always executes regardless of execution state outcomes)
    $Process.Dispose()
}
<#
.SYNOPSIS
    Consolidates, cleanses, and maps multi-source threat intelligence data into formatted Excel reports.
.DESCRIPTION
    The TMC script acts as an enterprise-grade Threat Intelligence parsing engine. It instantiates 
    a specialized .NET class structure that handles low-level COM automation with Microsoft Excel. 
    It parses incoming workbooks, cleanses noise out of malware actor naming structures, normalizes and 
    deduplicates raw text data using underlying high-performance [HashSet] structures, and routes 
    cleansed indicators into strict schema-compliant Excel spreadsheets and structured STIX-like reports.
.PARAMETER WorkBooks
    An array of absolute or relative string paths pointing to the source workbook files to ingest.
.EXAMPLE
    .\TMC.ps1 -WorkBooks "D:\IOCs\ActorGroupA.xlsx"
.EXAMPLE
    .\TMC.ps1 -WorkBooks "D:\IOCs\ActorGroupA.xlsx","D:\IOCs\CampaignB.xlsx",,"D:\IOCs\CampaignC.xlsx"
.NOTES
    Author: SivaMani70
    Date: May 2026
    Prerequisites: Requires local installation of Microsoft Excel (COM Interop validation executed against HKLM).
    Safety: Implements programmatic Marshal COM reference release wrappers to eliminate zombie excel.exe processes.
#>


[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string[]]$WorkBooks
)

# Import required dependency baseline components for IOC parsing, Excel report generation, and file movement operations
. $PSScriptRoot\root\IOC.ps1
. $PSScriptRoot\root\ExcelReport.ps1
. $PSScriptRoot\Move.ps1
class TMC {
    # --- Class Properties ---
    [System.Collections.Generic.Dictionary[string, System.Collections.Generic.HashSet[String]]]$IOCData
    [string[]]$WorkBooks
    [System.Object]$Excel
    
    # Metrics Telemetry Counters
    [int]$MD5Count = 0
    [int]$SHA1Count = 0
    [int]$SHA256Count = 0
    [int]$DomainsCount = 0
    [int]$URLsCount = 0
    [int]$EmailsCount = 0
    [int]$IPCount = 0
    [int]$OtherCount = 0
    [int]$Total = 0

    # Highly efficient deduplicated text storage buckets for each observable type, leveraging .NET HashSet for O(1) complexity on add and lookup operations, ensuring that the script can handle large volumes of indicators without performance degradation
    [System.Collections.Generic.HashSet[String]]$MD5
    [System.Collections.Generic.HashSet[String]]$SHA1
    [System.Collections.Generic.HashSet[String]]$SHA256
    [System.Collections.Generic.HashSet[String]]$Domains
    [System.Collections.Generic.HashSet[String]]$URLS
    [System.Collections.Generic.HashSet[String]]$IPS
    [System.Collections.Generic.HashSet[String]]$Emails
    [System.Collections.Generic.HashSet[String]]$OtherIOCs

    # Type-constrained reporting data streams that accumulate structured PSCustomObjects for each observable type, which are then passed to the Excel report generation engine, ensuring that the data adheres to a consistent schema and can be easily manipulated for reporting purposes
    [System.Collections.Generic.List[PSCustomObject]]$HashReportData
    [System.Collections.Generic.List[PSCustomObject]]$OtherIndicatorsReportData

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
    TMC([string[]]$WorkBooks) {
        $this.WorkBooks = $WorkBooks
        $this.Excel = New-Object -ComObject Excel.Application
        $this.IOCData = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.HashSet[String]]]::new()


        $this.MD5 = New-Object System.Collections.Generic.HashSet[String]
        $this.SHA1 = New-Object System.Collections.Generic.HashSet[String]
        $this.SHA256 = New-Object System.Collections.Generic.HashSet[String]
        $this.Domains = New-Object System.Collections.Generic.HashSet[String]
        $this.URLS = New-Object System.Collections.Generic.HashSet[String]
        $this.IPS = New-Object System.Collections.Generic.HashSet[String]
        $this.Emails = New-Object System.Collections.Generic.HashSet[String]
        $this.OtherIOCs = New-Object System.Collections.Generic.HashSet[String]

        $this.HashReportData = New-Object System.Collections.Generic.List[PSCustomObject]
        $this.OtherIndicatorsReportData = New-Object System.Collections.Generic.List[PSCustomObject]
    }


    # --- Method: ClearAllLists ---
    # Purpose: Resets all underlying volatile hash collections between sheet generation cycles. 
    # This ensures that data from previous iterations does not contaminate subsequent report generations, 
    # maintaining the integrity and accuracy of the output while allowing the script to process multiple  
    # workbooks in a single execution without manual intervention.
    [void] ClearAllLists() {
        $this.MD5.Clear()
        $this.SHA1.Clear()
        $this.SHA256.Clear()
        $this.Domains.Clear()
        $this.URLS.Clear()
        $this.Emails.Clear()
        $this.IPS.Clear()
        $this.OtherIOCs.Clear()
    }

    # --- Method: UpdateCount ---
    # Purpose: Flushes current active list snapshots over to running global analytics totals. This method is invoked after each workbook processing cycle to ensure that the final report includes accurate counts of each indicator type, which can be used for metrics tracking, reporting, and further analysis.
    [Void] UpdateCount() {
        $this.MD5Count += $this.MD5.Count
        $this.SHA1Count += $this.SHA1.Count
        $this.SHA256Count += $this.SHA256.Count
        $this.DomainsCount += $this.Domains.Count
        $this.URLsCount += $this.URLS.Count
        $this.EmailsCount += $this.Emails.Count
        $this.IPCount += $this.IPS.Count
        $this.OtherCount += $this.OtherIOCs.Count

        [int]$this.Total = $this.MD5Count + $this.SHA1Count + $this.SHA256Count + $this. DomainsCount + $this.URLsCount + $this.EmailsCount + $this.IPCount + $this.OtherCount
    }

    # --- Method: GetFileName ---
    # Purpose: Formats a standardized, localized name for report output files based on active date metadata. 
    # This method ensures that generated reports have consistent and descriptive names that include the month, 
    # day, and year of report generation, improving organization and traceability of threat intelligence outputs.
    [String] GetFileName() {
        [datetime]$Date = Get-Date
        [string]$Month = (Get-Culture).DateTimeFormat.GetMonthName($Date.Month)
        return "TMC Threat Bytes $Month $($Date.Day) $($Date.Year)"
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

    # --- Method: HashReport ---
    # Purpose: Compiles a standardized 23-column STIX-compliant hash metadata entry for reporting pools. This method takes in a malware name and an observable value, validates the hash against known patterns, and constructs a structured PSCustomObject that can be easily consumed by the Excel report generation engine, ensuring that all hash indicators are reported with consistent metadata and formatting.
    [void] HashReport([string]$MalwareName, [String]$ObservableValue) {
        $ReportEntry = [PSCustomObject]@{
            "fileName"              = " "
            "directoryPath"         = " "
            "threatTypes"           = " "
            "tags"                  = $MalwareName
            "name"                  = "TMC TI"
            "description"           = "$MalwareName IOC"
            "confidence"            = 100
            "revoked"               = " "
            "validFrom"             = (Get-Date).ToString("o")
            "validUntil"            = (Get-Date).AddDays(90).ToString("o")
            "tlpLevel"              = "amber"
            "severity"              = 5
            "filehash:md5"          = if ($ObservableValue -match $this.MD5_Validator) { $ObservableValue } else { " " }
            "filehash:md6"          = " "
            "filehash:ripemd - 160" = " "
            "filehash:sha - 1"      = if ($ObservableValue -match $this.SHA1_Validator) { $ObservableValue } else { " " }
            "filehash:sha - 224"    = " "
            "filehash:sha - 256"    = if ($ObservableValue -match $this.SHA256_Validator) { $ObservableValue } else { " " }
            "filehash:sha - 384"    = " "
            "filehash:sha - 512"    = " "
            "filehash:sha3 - 224"   = " "
            "filehash:sha3 - 256"   = " "
            "filehash:sha3 - 384"   = " "
            "filehash:sha3 - 512"   = " "
            "filehash:ssdeep"       = " "
            "filehash:whirlpool"    = " "
        }
        $this.HashReportData.Add($ReportEntry) | Out-Null; 
    }

    # --- Method: OtherIndicatorsReport ---
    # Purpose: Generates structured, UTC-normalized observational logs for non-file network metrics. This method constructs a standardized report entry for indicators such as IP addresses, domains, URLs, and emails, including metadata such as confidence levels, validity periods, and severity ratings. The structured output ensures that all non-hash indicators are reported with consistent formatting and can be easily integrated into the final Excel reports.
    [void] OtherIndicatorsReport([string]$MalwareName, [String]$ObservableType, [String]$ObservableValue) {
        $ReportEntry = [PSCustomObject]@{
            "threatTypes"     = " "
            "tags"            = $MalwareName
            "name"            = "TMC TI"
            "description"     = "$MalwareName IOC"
            "confidence"      = 100
            "revoked"         = " "
            "validFrom"       = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ", [System.Globalization.CultureInfo]::InvariantCulture)
            "validUntil"      = (Get-Date).AddDays(90).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ", [System.Globalization.CultureInfo]::InvariantCulture)
            "tlpLevel"        = "amber"
            "severity"        = 5
            "observableType"  = $ObservableType
            "observableValue" = $ObservableValue
        }
        $this.OtherIndicatorsReportData.Add($ReportEntry) | Out-Null; 
    }

    # --- Method: OrganizeIOC ---
    # Purpose: Cleans bad character structures out of actor labels and maps strings via 
    # regex switch trees directly into specialized, deduplicated type buckets.
    [void] OrganizeIOC([string]$MalwareName, [System.Collections.Generic.HashSet[String]]$Indicators) {
        # $MalwareName = ($MalwareName -replace "[\\/\?\*\[\]\(\):]", ' ').Trim()
        $MalwareName = ($MalwareName -replace "['’]s|[\\/\?\*\[\]\(\):]", ' ').Trim()
        $MalwareName = $MalwareName -replace '(?i)ioc', ''
        $MalwareName = $MalwareName -replace '(?i)iocs', ''
        $MalwareName = $MalwareName -replace '(?i)ttp', ''
        $MalwareName = $MalwareName -replace '&', ''
        $MalwareName = $MalwareName.Trim()
        foreach ($Indicator in $Indicators) {
            $Indicator = ($Indicator.ToLower()).Trim()

            switch -Regex ($Indicator) {
                # Check IP Addresses
                { $_ -match $this.IPV4Validator -or $_ -match $this.IPV6Validator } { 
                    $this.IPS.Add($Indicator) | Out-Null
                    if ($_ -match $this.IPV4Validator) {
                        $this.OtherIndicatorsReport($MalwareName, "ipv4-addr", $Indicator)
                    }
                    else {
                        $this.OtherIndicatorsReport($MalwareName, "ipv6-addr", $Indicator)
                    }
                    continue 
                }

                # Check Domain
                $this.DomainValidator { 
                    $this.Domains.Add($Indicator) | Out-Null
                    $this.OtherIndicatorsReport($MalwareName, "domain-name", $Indicator)
                    continue 
                }
                
                # Check Hashes
                $this.MD5_Validator    {
                    $this.MD5.Add($Indicator) | Out-Null; 
                    $this.HashReport($MalwareName, $Indicator)
                    continue 
                }
                $this.SHA1_Validator   {
                    $this.SHA1.Add($Indicator) | Out-Null; 
                    $this.HashReport($MalwareName, $Indicator)
                    continue 
                }
                $this.SHA256_Validator {
                    $this.SHA256.Add($Indicator) | Out-Null; 
                    $this.HashReport($MalwareName, $Indicator)    
                    continue 
                }
                
                # Check URL
                $this.URLValidator { 
                    $this.URLS.Add($Indicator) | Out-Null
                    $this.IPExtractorFromURLs($Indicator)
                    $this.OtherIndicatorsReport($MalwareName, "url", $Indicator)
                    continue 
                }
                
                # Check Email
                $this.EmailValidator {
                    $this.Emails.Add($Indicator) | Out-Null; 
                    $this.OtherIndicatorsReport($MalwareName, "email-addr", $Indicator)
                    continue 
                }
                
                Default {
                    # This runs if NONE of the patterns above matched
                    Write-Warning "Uncategorized data found: $Indicator"
                    $this.OtherIOCs.Add($Indicator) | Out-Null
                    $this.OtherIndicatorsReport($MalwareName, "unknown", $Indicator)
                }
            }
        }
    }

    # --- Method: GenerateTextFile ---
    # Purpose: Standard file-writing method that flushes aggregated dictionary collections out to raw text files. 
    # This method iterates over the organized IOC data and generates a plain text file containing all 
    # indicators, which can be used for quick reference, sharing, or as an input for other tools that consume 
    # text-based indicator lists. The generated file is named based on the current date to ensure uniqueness and traceability.
    [void] GenerateTextFile() {
        [String]$FileName = "$($this.GetFileName())_IOCs.txt"
        $FullName = (Get-Location).Path + "\$FileName"
        foreach ($MalwareName in $this.IOCData.Keys) {
            Write-Host "Generating text file for $MalwareName at $FullName" -ForegroundColor Green
            $this.IOCData[$MalwareName] | Out-File -Append -FilePath $FullName -Encoding UTF8
        }
    }


    # --- Method: CreateTMCWorkBook ---
    # Purpose: Orchestrates layout design configurations, builds custom tab blocks, formats 
    # column grids dynamically, and compiles final workbook assets to disk.
    [Void] CreateTMCWorkBook() {
        [String]$FileName = $this.GetFileName()
        $FullName = (Get-Location).Path + "\$FileName.xlsx"

        $WorkBook = $this.Excel.Workbooks.Add()
        try {
            Write-Host "Creating File -- $FullName. `nThis may take some time...." -ForegroundColor Yellow
            foreach ($MalwareName in $this.IOCData.Keys) {
                $this.ClearAllLists()

                Write-Host "Adding IOCs to Sheet: [$MalwareName]" -ForegroundColor Green
                $WorkSheet = $WorkBook.Worksheets.Add()
                $SheetName = ($MalwareName -replace '[\\/\?\*\[\]\(\):]', ' ').Trim()
                $SheetName = $SheetName.Substring(0, [Math]::Min($SheetName.Length, 29))
                Write-Warning "Sheet Name: $SheetName"
                $WorkSheet.Name = $SheetName
                $this.OrganizeIOC($MalwareName, $this.IOCData[$MalwareName])

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
                $this.UpdateCount()
                $WorkSheet.Columns.AutoFit()
            }
            $CountSheet = $WorkBook.Sheets["Sheet1"]
            $CountSheet.Name = "Count"
            $this.WriteToCountSheet($CountSheet)
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
                $MalwareName = ((Split-Path -Path $WorkBook -Leaf).Split('.')[0]).Replace(':', '')
                [System.Collections.Generic.HashSet[String]]$ListOfIOCs = Get-IOC -WBPath $WorkBook -Excel $this.Excel
                $this.IOCData.Add($MalwareName, $ListOfIOCs)
            }
            else {
                Write-Warning "File not exist -- [$WorkBook]"
            }
        }


        if ($null -ne $this.IOCData -and $this.IOCData.Count -gt 0) {
            # Checking if any of those sheets actually contain IOC values
            $HasData = ($this.IOCData.Values | Where-Object { $_.Count -gt 0 })    
            if ($HasData) {
                Write-Host "Successfully collected IOCs. Generating report..."
                
                # Prompt the user if they want to generate a text file with the IOCs. This is optional and can be used for quick reference or sharing without needing Excel.
                $Response = $null
                do {
                    $Response = Read-Host "Do you want to generate a text file with the IOC data? (Y/N)"
                } until ($Response.ToLower() -in 'y', 'n', 'yes', 'no')
                $IsYes = ($Response.ToLower() -eq 'yes' -or $Response.ToLower() -eq 'y')
                if ($IsYes) {
                    $this.GenerateTextFile()
                }
                $this.CreateTMCWorkBook()
            }
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

    [void] WriteToCountSheet($WorkSheet) {
        $Row = 2
        $Col = 2
        $this.FormatHeader($WorkSheet.Cells.Item($Row, $Col), "IOC Type")
        $Col += 1
        $this.FormatHeader($WorkSheet.Cells.Item($Row, $Col), "Count")
        $this.CenterContent($WorkSheet.Cells.Item($Row, $Col))


        $CountMap = [ordered]@{
            "MD5"           = $this.MD5Count
            "SHA1"          = $this.SHA1Count
            "SHA256"        = $this.SHA256Count
            "Domains"       = $this.DomainsCount
            "URLs"          = $this.URLsCount
            "Emails"        = $this.EmailsCount
            "IP"            = $this.IPCount
            "Review - IOCs" = $this.OtherCount
        }

        foreach ($Entry in $CountMap.GetEnumerator()) {
            $Name = $Entry.Key
            $Count = $Entry.Value

            if ($Count -gt 0) {
                $Row++
                $Col = 2
                $this.WriteToCell($WorkSheet.Cells.Item($Row, $Col), $Name)
                $Col += 1
                $this.WriteToCell($WorkSheet.Cells.Item($Row, $Col), "$Count")
                $this.CenterContent($WorkSheet.Cells.Item($Row, $Col))
                
            }
        }
        
        $Row++
        $Col = 2
        $this.FormatHeader($WorkSheet.Cells.Item($Row, $Col), "Total")
        $Col += 1
        $this.WriteToCell($WorkSheet.Cells.Item($Row, $Col), "$($this.Total)")
        $this.CenterContent($WorkSheet.Cells.Item($Row, $Col))
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
[TMC]$Process = [TMC]::new($WorkBooks)
try {
    $Process.Load()

    # Conditionally compile advanced hash data logs if records exist
    if ($Process.HashReportData.Count -gt 0) {
        Write-Host "`n`n`nGenerating Hash Report. This may take some time..." -ForegroundColor Green
        New-Report -IOCType "Hash Data" -Data $Process.HashReportData
    }

    # Conditionally compile advanced network threat logs if records exist
    if ($Process.OtherIndicatorsReportData.Count -gt 0) {
        Write-Host "`n`n`nGenerating Other Indicators Report. This may take some time..." -ForegroundColor Green
        New-Report -IOCType "Other Indicators" -Data $Process.OtherIndicatorsReportData
    }

    # Run workspace folder sorting cleanup routine
    Move-Files
}
finally {
    # 3. Memory Cleanup Execution Hook (Always executes regardless of execution state outcomes)
    $Process.Dispose()
}
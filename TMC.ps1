[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string[]]$WorkBooks
)

. $PSScriptRoot\root\IOC.ps1
. $PSScriptRoot\root\ExcelReport.ps1
. $PSScriptRoot\Move.ps1
class TMC {
    [System.Collections.Generic.Dictionary[string, System.Collections.Generic.HashSet[String]]]$IOCData
    [string[]]$WorkBooks
    [System.Object]$Excel

    [int]$MD5Count = 0
    [int]$SHA1Count = 0
    [int]$SHA256Count = 0
    [int]$DomainsCount = 0
    [int]$URLsCount = 0
    [int]$EmailsCount = 0
    [int]$IPCount = 0
    [int]$OtherCount = 0
    [int]$Total = 0
    
    [System.Collections.Generic.HashSet[String]]$MD5
    [System.Collections.Generic.HashSet[String]]$SHA1
    [System.Collections.Generic.HashSet[String]]$SHA256
    [System.Collections.Generic.HashSet[String]]$Domains
    [System.Collections.Generic.HashSet[String]]$URLS
    [System.Collections.Generic.HashSet[String]]$IPS
    [System.Collections.Generic.HashSet[String]]$Emails
    [System.Collections.Generic.HashSet[String]]$OtherIOCs

    [System.Collections.Generic.List[PSCustomObject]]$HashReportData
    [System.Collections.Generic.List[PSCustomObject]]$OtherIndicatorsReportData

    [String]$MD5_Validator = "^[a-fA-F0-9]{32}$"
    [String]$SHA1_Validator = "^[a-fA-F0-9]{40}$"
    [String]$SHA256_Validator = "^[a-fA-F0-9]{64}$"
    [String]$DomainValidator = "^(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})?$"
    [String]$URLValidator = "^(https?|hxxps?|ftp):\/\/[^\s/$.?#].[^\s]*$"
    [String]$EmailValidator = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    [String]$IPV4Validator = "^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
    [String]$IPV6Validator = "^(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))$"


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

    [String] GetFileName() {
        [datetime]$Date = Get-Date
        [string]$Month = (Get-Culture).DateTimeFormat.GetMonthName($Date.Month)
        return "TMC Threat Bytes $Month $($Date.Day) $($Date.Year)"
    }

    [bool] IsValidPath([string]$Path) {
        return Test-Path -Path $Path
    }

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

    [void] GenerateTextFile() {
        [String]$FileName = "$($this.GetFileName())_IOCs.txt"
        $FullName = (Get-Location).Path + "\$FileName"
        foreach ($MalwareName in $this.IOCData.Keys) {
            Write-Host "Generating text file for $MalwareName at $FullName" -ForegroundColor Green
            $this.IOCData[$MalwareName] | Out-File -Append -FilePath $FullName -Encoding UTF8
        }
    }


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
            if ($null -ne $WorkBook) {
                # Set DisplayAlerts to false so it doesn't pop up a "Save changes?" window if it crashes
                $this.Excel.DisplayAlerts = $false
                $WorkBook.Close($false) # False means don't save pending changes on a crash
                $this.Excel.DisplayAlerts = $true
            }
        }
    }


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
        $Cell.Interior.ColorIndex = 37 # Light Blue
        $Cell.Font.Bold = $true
    }
    
    [void] CenterContent($Cell) {
        $Cell.HorizontalAlignment = 3 # Center
    }

    [void] WriteToCell($Cell, $Value) {
        $Cell.Value = $Value
        $Cell.Borders.LineStyle = 1
        $Cell.Borders.ColorIndex = 1
    }

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

# Script Execution starts from the below
if (-not (Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Office\*\Excel")) {
    Write-Error "No Excel Module found in the System"    
    return
}
[TMC]$Process = [TMC]::new($WorkBooks)
try {
    $Process.Load()

    if ($Process.HashReportData.Count -gt 0) {
        Write-Host "`n`n`nGenerating Hash Report. This may take some time..." -ForegroundColor Green
        New-Report -IOCType "Hash Data" -Data $Process.HashReportData
    }

    if ($Process.OtherIndicatorsReportData.Count -gt 0) {
        Write-Host "`n`n`nGenerating Other Indicators Report. This may take some time..." -ForegroundColor Green
        New-Report -IOCType "Other Indicators" -Data $Process.OtherIndicatorsReportData
    }
    Move-Files
}
finally {
    $Process.Dispose()
}
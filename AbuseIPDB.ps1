<#
.SYNOPSIS
    Checks IP address reputations against the AbuseIPDB API.
.DESCRIPTION
    This script extracts IPv4 and IPv6 addresses from a target raw file, de-duplicates 
    them using a .NET HashSet, evaluates their reputation via the AbuseIPDB v2 API endpoint, 
    and pipes the structured results into an Excel report generation tool.
.PARAMETER APIKEY
    The authorization credential key required to authenticate to the AbuseIPDB API.
.PARAMETER FilePath
    The system file path pointing to the text artifact containing raw IOC data.
.EXAMPLE
    .\AbuseIPDB.ps1 -APIKEY "AbuseIPDB_Secret_Token_123" -FilePath "C:\Threats\iocs.txt"
.EXAMPLE
    .\AbuseIPDB.ps1 -APIKEY "AbuseIPDB API Key" -FilePath "C:\Threats\ips.txt"
.NOTES
    Author: SivaMani70
    Date: May 2026
    Dependency: Requires external scripts 'root\IOC.ps1' and 'root\ExcelOrCSVReport.ps1'.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$APIKEY,
    [Parameter(Mandatory)]
    [string]$FilePath
)

# Import required baseline functional scripts
. $PSScriptRoot\root\IOC.ps1
. $PSScriptRoot\root\ExcelOrCSVReport.ps1

class AbuseIPRep {
    # Validations & Constant Mappings
    [string]$IPV4Validator = "^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
    [String]$IPV6Validator = "^(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))$"
    [string]$ENDPOINT = "https://api.abuseipdb.com/api/v2/check"
    [string]$Type = "IP"

    # State Variables
    [string]$FilePath
    [System.Collections.Generic.HashSet[String]]$ListOfIP
    [System.Collections.Generic.List[PSCustomObject]]$Data
    [Hashtable]$Headers = @{}
    [bool]$RateLimitHit

    # --- Constructor ---
    # Initializes properties, prepares telemetry headers, and configures memory spaces
    AbuseIPRep([string]$Path, [string]$Key) {
        $this.FilePath = $Path
        $this.RateLimitHit = $false

        $this.ListOfIP = New-Object System.Collections.Generic.HashSet[String]
        $this.Data = New-Object System.Collections.Generic.List[PSCustomObject]
        $this.Headers.Add("accept", "application/json")
        $this.Headers.Add("key", $Key)
    }

    # --- Method: ExtractIP ---
    # Purpose: Validates file existence, invokes the file parser helper, and filters lines 
    # against rigorous Regex patterns to capture true IPv4/IPv6 indicators into a unique set.
    [void] ExtractIP() {
        if (-not (Test-Path -Path $this.FilePath)) {
            Write-Error "File $($this.FilePath) is Invalid/File Not exist"
            return

        }
        Write-Host "Extracting IPS"
        [System.Collections.Generic.HashSet[String]]$IOCs = Get-IOC_TXT -FilePath $this.FilePath
        foreach ($IOC in $IOCs) {
            $IOC = $IOC.Trim()
            if (($IOC -match $this.IPV4Validator) -or ($IOC -match $this.IPV6Validator)) {
                # HashSet inherently enforces uniqueness by ignoring pre-existing strings, so no need for explicit duplicate checks
                $this.ListOfIP.Add($IOC) | Out-Null
            }
        }
    }

    # --- Method: Check ---
    # Purpose: Orchestrates extraction, loops over validated targets to perform REST calls, 
    # processes HTTP responses into clean PSCustomObjects, handles status limits, 
    # and fires final reporting engine components.
    [void] Check() {
        $this.ExtractIP()
        Write-Host "$($this.ListOfIP.Count) - IP(s) found in the file $($this.FilePath)"
        if ($this.ListOfIP.Count -eq 0) { return }
        Write-Host "Checking IP reputation..." -ForegroundColor Green

        foreach ($Ip in $this.ListOfIP) {

            if ($this.RateLimitHit) {
                Write-Warning "AbuseIPDB API: Rate limit has been hit. Stopping further requests."
                break
            }
            
            Write-Host "IP: $Ip" -ForegroundColor Green

            [Hashtable]$QueryParameters = @{
                "ipAddress" = $Ip
                "verbose"   = " "
            }

            try {
                $Response = Invoke-RestMethod -Method Get -Uri $this.ENDPOINT -Headers $this.Headers -Body $QueryParameters
                $D = $Response.data

                # Appends successful API elements into a generic collection payload for later report generation
                $this.Data.Add([PSCustomObject]@{
                        IPAddress            = $D.ipAddress
                        ISP                  = $D.isp
                        TotalReports         = $D.totalReports
                        AbuseConfidenceScore = $D.abuseConfidenceScore
                        Domain               = $D.domain
                        Whitelisted          = $D.isWhitelisted
                        IsTor                = $D.isTor
                        UsageType            = $D.usageType
                        CountryCode          = $D.countryCode
                        CountryName          = $D.countryName
                    })

            }
            catch {
                $Exception = $_.Exception
                $ErrorMessage = $_.Exception.Message
                $Resp = $exception.Response

                if ($null -ne $Resp) {
                    $StatusCode = [int]$Resp.StatusCode
                    # REST Exception handling maps API error status boundaries to user-friendly warnings and logs for operational visibility
                    switch ($StatusCode) {
                        429 {
                            $this.RateLimitHit = $true
                            Write-Warning "AbuseIPDB API: Rate limit exceeded (429). Please wait before retrying."
                            break
                        }
                        Default {
                            Write-Warning "AbuseIPDB API: Received error code $StatusCode"
                            Write-Error "AbuseIPDB API: Received error Message $ErrorMessage while handling $Ip"
                        }
                    }
                }
            }
        }

        # Validate that database entities were successfully parsed before calling report hooks, ensuring that empty datasets do not trigger redundant report generation processes
        if ($this.Data.Count -le 0) {
            Write-Warning "No data captured to create a report"
            return
        }
        # Passes results to reporting function that transforms the structured collection into an Excel report format, abstracting away presentation logic from data processing
        New-Report -IOCType $this.Type -Data $this.Data
    }
}

([AbuseIPRep]::new($FilePath, $APIKEY)).Check()
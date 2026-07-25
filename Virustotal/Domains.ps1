<#
.SYNOPSIS
    Backend worker script for retrieving VirusTotal reputation scores for Domains.

.DESCRIPTION
    This script defines and executes the `DomainRep` class, which interacts with the 
    VirusTotal API v3. It extracts valid domains from a specified file, and queries their reputation. 

    If the domain is not found in VirusTotal's database, it will log a warning with 404. The script also handles API rate limits gracefully.

    NOTE: This script is not intended to be executed directly by the user. It acts as 
    a specialized module invoked by wrapper scripts like MIXR.ps1 and SOLO.ps1.

.PARAMETER APIKEY
    The VirusTotal API v3 key required for authentication. This is passed down 
    dynamically from the invoking wrapper script.

.PARAMETER FilePath
    The absolute or relative path to the text file containing the Indicators of 
    Compromise (IOCs) to be processed. Passed down from the wrapper script.

.EXAMPLE
    # Invocation via the MIXR wrapper script:
    .\MIXR.ps1 -APIKEY "VT_Crypto_Token_XYZ" -IOC_FilePath "C:\Threats\today_iocs.txt"

.EXAMPLE
    # Invocation via the SOLO wrapper script:
    .\SOLO.ps1 -APIKEY "VT_Crypto_Token_XYZ" -IOC_FilePath "C:\Threats\today_iocs.txt"

.NOTES
    Author: SivaMani70
    Date: May 2026

    Dependencies: Requires `IOC.ps1` to be present in the `..\root\` directory relative 
    to this script's location for the `Get-IOC_TXT` function.
#>


[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$APIKEY,
    [Parameter(Mandatory)]
    [string]$FilePath
)

. $PSScriptRoot\..\root\IOC.ps1
Class DomainRep {
    [String]$DomainValidator = "^(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})?$"
    [string]$ENDPOINT = "https://www.virustotal.com/api/v3/domains/"

    [string]$FilePath
    [System.Collections.Generic.HashSet[String]]$ListOfDomains
    [System.Collections.Generic.List[PSCustomObject]]$Data
    [Hashtable]$Headers = @{}
    [bool]$RateLimitHit


    DomainRep([string]$Path, [string]$Key) {
        $this.FilePath = $Path
        $this.RateLimitHit = $false

        $this.ListOfDomains = New-Object System.Collections.Generic.HashSet[String]
        $this.Data = New-Object System.Collections.Generic.List[PSCustomObject]
        $this.Headers.Add("accept", "application/json")
        $this.Headers.Add("x-apikey", $Key)
    }

    [void] ExtractDomains() {
        if (-not (Test-Path -Path $this.FilePath)) {
            Write-Error "File $($this.FilePath) is Invalid/File Not exist"
            return

        }
        Write-Host "Extracting Domains"
        [System.Collections.Generic.HashSet[String]]$IOCs = Get-IOC_TXT -FilePath $this.FilePath
        foreach ($IOC in $IOCs) {
            $IOC = $IOC.Trim()
            if ($IOC -match $this.DomainValidator) {
                $this.ListOfDomains.Add($IOC) | Out-Null
            }
        }
    }


    [System.Collections.Generic.List[PSCustomObject]] Check() {
        $this.ExtractDomains()
        Write-Host "$($this.ListOfDomains.Count) - Domain(s) found in the file $($this.FilePath)"
        if ($this.ListOfDomains.Count -eq 0) { return $this.Data }
        Write-Host "Checking Domain reputation..." -ForegroundColor Green

        foreach ($Domain in $this.ListOfDomains) {
            if ($this.RateLimitHit) {
                Write-Warning "VirusTotal API: Rate limit has been hit. Stopping further requests."
                break
            }

            Write-Host "Domain: $Domain" -ForegroundColor Green
            [String]$FinalURL = $this.ENDPOINT + $Domain
            try {
                $Response = Invoke-RestMethod -Method Get -Uri $FinalURL -Headers $this.Headers
                $D = $Response.data
                $Domain = $D.id 
                $Harmless = $D.attributes.last_analysis_stats.harmless
                $Malicious = $D.attributes.last_analysis_stats.malicious
                $Suspicious = $D.attributes.last_analysis_stats.suspicious
                $Undetected = $D.attributes.last_analysis_stats.undetected
                [int]$Total = $Harmless + $Malicious + $Suspicious + $Undetected
                [int]$MaliciousCount = $Malicious + $Suspicious

                $this.Data.Add([PSCustomObject]@{
                        Domain       = $Domain
                        Result       = "$MaliciousCount//$Total"
                        TotalChecked = $Total
                        Harmless     = $Harmless
                        Malicious    = $Malicious
                        Suspicious   = $Suspicious
                        Undetected   = $Undetected
                    })

            }
            catch {
                $Exception = $_.Exception
                $ErrorMessage = $_.Exception.Message
                $Resp = $exception.Response

                if ($null -ne $Resp) {
                    $StatusCode = [int]$Resp.StatusCode

                    switch ($StatusCode) {
                        429 {
                            $this.RateLimitHit = $true
                            Write-Warning "VirusTotal API: Rate limit exceeded (429). Please wait before retrying."
                            break
                        }
                        Default {
                            Write-Warning "VirusTotal API: Received error code $StatusCode"
                            Write-Error "VirusTotal API: Received error Message $ErrorMessage while handling $Domain"
                        }
                    }
                }
            }
        }

        if ($this.Data.Count -le 0) {
            Write-Warning "No data captured to create a report"
        }
        # New-Report -IOCType $this.Type -Data $this.Data
        return $this.Data
    }
}

([DomainRep]::new($FilePath, $APIKEY)).Check()
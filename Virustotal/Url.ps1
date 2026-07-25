<#
.SYNOPSIS
    Backend worker script for retrieving VirusTotal reputation scores for URLs.

.DESCRIPTION
    This script defines and executes the `UrlRep` class, which interacts with the 
    VirusTotal API v3. It extracts valid URLs from a specified file, converts them 
    into VT's required Base64 format, and queries their reputation. 

    If a URL is unknown to VirusTotal, the script automatically submits it for scanning 
    and implements a polling mechanism (retrying every 10 seconds for up to 2 minutes) 
    to wait for the asynchronous analysis to complete.

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

# Dot-source the IOC extraction dependency relative to the script's root directory
. $PSScriptRoot\..\root\IOC.ps1

class UrlRep {
    [String]$URLValidator = "^(https?|hxxps?|ftp):\/\/[^\s/$.?#].[^\s]*$"
    [string]$ScanEndpoint = "https://www.virustotal.com/api/v3/urls"

    [string]$FilePath
    [System.Collections.Generic.HashSet[String]]$ListOfUrls
    [System.Collections.Generic.List[PSCustomObject]]$Data
    [System.Collections.Generic.HashSet[String]]$SelfEndpoints
    [Hashtable]$Headers = @{}
    [bool]$RateLimitHit

    # Constructor: Initializes collections and prepares HTTP headers
    UrlRep([string]$Path, [string]$Key) {
        $this.FilePath = $Path
        $this.RateLimitHit = $false

        $this.ListOfUrls = New-Object System.Collections.Generic.HashSet[String]
        $this.Data = New-Object System.Collections.Generic.List[PSCustomObject]
        $this.SelfEndpoints = New-Object System.Collections.Generic.HashSet[String]

        # Build VirusTotal API v3 required headers
        $this.Headers.Add("accept", "application/json")
        $this.Headers.Add("x-apikey", $Key)
        $this.Headers.Add("content-type", "application/x-www-form-urlencoded")
    }

    # Reads the target file, extracts valid URLs using Regex, and stores them in a HashSet
    [void] ExtractUrls() {
        if (-not (Test-Path -Path $this.FilePath)) {
            Write-Error "File $($this.FilePath) is Invalid/File Not exist"
            return

        }
        Write-Host "Extracting Urls"
        [System.Collections.Generic.HashSet[String]]$IOCs = Get-IOC_TXT -FilePath $this.FilePath
        foreach ($IOC in $IOCs) {
            $IOC = $IOC.Trim()
            if (($IOC -match $this.URLValidator)) {
                $CleanUrl = $IOC -replace '(?i)^hxxp', 'http'
                $this.ListOfUrls.Add($CleanUrl) | Out-Null
            }
        }
    }

    # Centralized error handler for HTTP web exceptions (e.g., 404, 429)
    [void] HandleWebException([object]$ExceptionObject, [string]$Url) {
        $Exception = $ExceptionObject.Exception
        $ErrorMessage = $ExceptionObject.Exception.Message
        $Resp = $Exception.Response

        if ($null -ne $Resp) {
            $StatusCode = [int]$Resp.StatusCode

            switch ($StatusCode) {
                404 {
                    Write-Warning "VirusTotal API: Endpoint not found (404). The URL may not be recognized by VirusTotal."
                    break
                }
                429 {
                    $this.RateLimitHit = $true
                    Write-Warning "VirusTotal API: Rate limit exceeded (429). Please wait before retrying."
                    break
                }
                Default {
                    Write-Warning "VirusTotal API: Received error code $StatusCode"
                    Write-Error "VirusTotal API: Received error Message $ErrorMessage while handling $Url"
                }
            }
        }
    }
    

    # Converts standard URL to VirusTotal's required Base64url format (un-padded, URL-safe)
    [string] GetBase64UrlId([string]$Url) {
        $Bytes = [System.Text.Encoding]::UTF8.GetBytes($Url)
        return [Convert]::ToBase64String($Bytes).Replace('+', '-').Replace('/', '_').TrimEnd('=')
    }

    # Submits a new/unknown URL to VirusTotal for active scanning and captures the queue status link
    [void] GetSelfLink($Url) {
        $ContentType = "application/x-www-form-urlencoded"
        $EncodedUrl = "url=$([System.Uri]::EscapeDataString($Url))"
        try {
            Write-Host "Submitting URL for Self-link and analysis: $Url" -ForegroundColor Green
            $Response = Invoke-RestMethod -Method Post -Uri $this.ScanEndpoint -Headers $this.Headers -Body $EncodedUrl -ContentType $ContentType
            if ($Response.data) {
                $D = $Response.data
                $this.SelfEndpoints.Add($D.links.self) | Out-Null
            }
        }
        catch {
            $this.HandleWebException($_, $Url)
        }
    }

    # Checks VT for existing URL reports. If it gets a 404, it triggers an active scan via GetSelfLink()
    [void] GetResponse($Url) {
        $Base64UrlId = $this.GetBase64UrlId($Url)
        $Base64IDEndpoint = "$($this.ScanEndpoint)/$Base64UrlId"
        $SanitizedUrl = $Url.Replace(":", "[:]").Replace(".", "[.]")
        try {
            $Response = Invoke-RestMethod -Method Get -Uri $Base64IDEndpoint -Headers $this.Headers

            if ($null -ne $Response -and $null -ne $Response.data) {
                Write-Host "`nURL found in VirusTotal database. Retrieving analysis report for: $SanitizedUrl" -ForegroundColor Green
                $D = $Response.data
                $this.ParseImmediateReport($D)
            }
            else {
                Write-Host "Unable to parse the Response data. Submitting URL for re analysis: $SanitizedUrl" -ForegroundColor Yellow
                $this.GetSelfLink($Url)
            }
        }
        catch {
            $Exception = $_.Exception
            $Resp = $Exception.Response
            if ($null -ne $Resp) {
                $StatusCode = [int]$Resp.StatusCode
                if ($StatusCode -eq 404) {
                    Write-Host "`nURL not found in VirusTotal database. Submitting for analysis: $SanitizedUrl" -ForegroundColor Yellow
                    $this.GetSelfLink($Url)
                }
                else {
                    $this.HandleWebException($_, $Url)
                }
            }
        }
    }
    
    # Loops through all identified URLs to query them while managing rate limits
    # Gets the analysis report for each url, if the url is not found in the VirusTotal database, it will submit the url for analysis and capture the self-link endpoint for subsequent status checks.
    [void] ProcessUrlBatch() {
        foreach ($Url in $this.ListOfUrls) {
            if ($this.RateLimitHit) { 
                Write-Warning "Rate limit hit. Stopping further requests to VirusTotal API."    
                break 
            }
            $SanitizedUrl = $Url.Replace(":", "[:]").Replace(".", "[.]")
            Write-Host "`nUrl: $SanitizedUrl" -ForegroundColor Yellow
            $this.GetResponse($Url);
        }
    }

    # If the Virustotal Database have the url report it will extract the relevant information and add it to the Data list.
    # Parses immediate report JSON and maps it to a standard PowerShell custom object
    [void] ParseImmediateReport($D) {
        $Url = ($D.attributes.url).Replace(":", "[:]").Replace(".", "[.]")
        $Harmless = $D.attributes.last_analysis_stats.harmless
        $Malicious = $D.attributes.last_analysis_stats.malicious
        $Suspicious = $D.attributes.last_analysis_stats.suspicious
        $Undetected = $D.attributes.last_analysis_stats.undetected
        [int]$Total = $Harmless + $Malicious + $Suspicious + $Undetected
        [int]$MaliciousCount = $Malicious + $Suspicious

        $this.Data.Add([PSCustomObject]@{
                URL          = $Url
                Result       = "$MaliciousCount//$Total"
                TotalChecked = $Total
                Harmless     = $Harmless
                Malicious    = $Malicious
                Suspicious   = $Suspicious
                Undetected   = $Undetected
            })
    }

    # Checks and process the data if the analysis status is completed
    # Parses queued analysis JSON (has a different structure than immediate reports)
    [void] ParseQueuedReport($D) {
        $Url = ($D.attributes.url).Replace(":", "[:]").Replace(".", "[.]")
        $Harmless = $D.attributes.stats.harmless
        $Malicious = $D.attributes.stats.malicious
        $Suspicious = $D.attributes.stats.suspicious
        $Undetected = $D.attributes.stats.undetected
        [int]$Total = $Harmless + $Malicious + $Suspicious + $Undetected
        [int]$MaliciousCount = $Malicious + $Suspicious

        $this.Data.Add([PSCustomObject]@{
                URL          = $Url
                Result       = "$MaliciousCount//$Total"
                TotalChecked = $Total
                Harmless     = $Harmless
                Malicious    = $Malicious
                Suspicious   = $Suspicious
                Undetected   = $Undetected
            })
    }

    # Checks and process the data if the analysis status is completed
    [System.Object] GetCompletionReport([string]$Endpoint) {
        $Response = Invoke-RestMethod -Method Get -Uri $Endpoint -Headers $this.Headers
        $D = $Response.data
        
        if ($D.attributes.status -eq "completed") {
            Write-Host "Analysis completed for endpoint: $Endpoint" -ForegroundColor Green
            $this.ParseQueuedReport($D)
        }
        return $D
    }

    # If the analysis status is "in-progress" but there are actionable data available, this method will extract the relevant information and add it to the Data list. It also logs the status and endpoint for reference.
    [bool] ProcessActionableData($D) {
        Write-Host "Scanning Status is in progress, but we have some analysis data available. Proceeding with the completed IOCs responses..." -ForegroundColor Green
        Write-Host "`n`n`nScanning Status: $($D.attributes.status)" -ForegroundColor Yellow
        $this.ParseQueuedReport($D)
        return $true
    }

    # This method handles the retry logic for endpoints that are still being analyzed by VirusTotal. It checks the status of each endpoint and waits for the analysis to complete, retrying every 10 seconds until a completed status is received or the maximum number of retries is reached.
    [System.Collections.Generic.HashSet[String]] RetryEndpoints([System.Collections.Generic.HashSet[String]]$AwaitingEndpoints) {
        Write-Host "VirusTotal can take up to 2 minutes to analyze a Url, so we will retry every 10 seconds until we get a completed status or we reach the max retries" -ForegroundColor Yellow

        [System.Collections.Generic.HashSet[String]]$FailedEndpoints = New-Object System.Collections.Generic.HashSet[String]
        foreach ($Endpoint in $AwaitingEndpoints) {
            if ($this.RateLimitHit) { 
                Write-Warning "Rate limit hit. Stopping further requests to VirusTotal API."    
                break 
            }
            try {
                $D = $this.GetCompletionReport($Endpoint)
                $IsReady = ($D.attributes.status -eq "completed")
                if ($IsReady) {
                    continue
                }

                
                $RetryCount = 1
                $MaxRetries = 20
                $HasActionableData = ($null -ne $D.attributes.stats -and (($D.attributes.stats.harmless + $D.attributes.stats.malicious + $D.attributes.stats.suspicious + $D.attributes.stats.undetected) -gt 0))

                if ($D.attributes.status -eq "completed" -or ($D.attributes.status -eq "in-progress" -and $HasActionableData)) {
                    $IsReady = $this.ProcessActionableData($D)
                }
                else {
                    while (-not $IsReady -and $RetryCount -le $MaxRetries) {
                        Write-Host "Analysis Status: $($D.attributes.status) ... (Attempt: $RetryCount/$MaxRetries)" -ForegroundColor Yellow
                        Start-Sleep -Seconds 10
                        $D = $this.GetCompletionReport($Endpoint)

                        $IsReady = ($D.attributes.status -eq "completed")
                        if ($IsReady) {
                            break
                        }
                        $HasActionableData = ($null -ne $D.attributes.stats -and (($D.attributes.stats.harmless + $D.attributes.stats.malicious + $D.attributes.stats.suspicious + $D.attributes.stats.undetected) -gt 0))
                        if ($D.attributes.status -eq "in-progress" -and $HasActionableData) {
                            $IsReady = $this.ProcessActionableData($D)                        
                        }
                        $RetryCount++
                    }
                }

                if (-not $IsReady) {
                    Write-Warning "Unable to get the report for the endpoint: $Endpoint after 120 seconds."
                    $FailedEndpoints.Add($Endpoint) | Out-Null
                    continue 
                }
            }
            catch {
                $this.HandleWebException($_, $Endpoint)
            }
        }
        return $FailedEndpoints
    }

    # Execution of the main logic to check the Urls in the file and get their analysis report from VirusTotal.
    # Main orchestrator method to execute the class lifecycle
    [System.Collections.Generic.List[PSCustomObject]] Check() {
        $this.ExtractUrls()
        Write-Host "$($this.ListOfUrls.Count) - Url(s) found in the file $($this.FilePath)"
        if ($this.ListOfUrls.Count -eq 0) { return $this.Data }
        Write-Host "Checking Url reputation..." -ForegroundColor Green
        [System.Collections.Generic.HashSet[String]]$AwaitingEndpoints = New-Object System.Collections.Generic.HashSet[String]
        $this.ProcessUrlBatch()

        foreach ($Endpoint in $this.SelfEndpoints) {

            if ($this.RateLimitHit) { 
                Write-Warning "Rate limit hit. Stopping further requests to VirusTotal API."    
                break 
            }

            try {
                $D = $this.GetCompletionReport($Endpoint)
                if ($D.attributes.status -ne "completed") {
                    Write-Host "`n`n`nScanning Status: $($D.attributes.status)" -ForegroundColor Yellow
                    Write-Host "Endpoint: $Endpoint" -ForegroundColor Yellow
                    Write-Host "Waiting for the analysis to complete..." -ForegroundColor Yellow
                    Write-Host "Proceeding with the completed IOCs responses..." -ForegroundColor Green
                    $AwaitingEndpoints.Add($Endpoint) | Out-Null
                }
            }
            catch {
                $this.HandleWebException($_, $Endpoint)
            }
        }
        

        if ($AwaitingEndpoints.Count -gt 0) {
            Write-Warning "The following endpoints are still being analyzed:"
            $AwaitingEndpoints | ForEach-Object { Write-Warning $_ }
            $FailedEndpoints = $this.RetryEndpoints($AwaitingEndpoints)

            if ($FailedEndpoints.Count -gt 0) {
                Write-Warning "The following endpoints could not be processed after retries:"
                $FailedEndpoints | ForEach-Object { Write-Warning $_ }
                Write-Host "`n`nSaving the failed endpoints to Awaiting_Endpoints.txt for future reference." -ForegroundColor Yellow
                $AwaitingFileName = "Awaiting_Endpoints.txt"
                $FailedEndpoints | Out-File -FilePath "$((Get-Location).Path)\$AwaitingFileName" -Encoding utf8
            }
            else {
                Write-Host "All endpoints processed successfully after retries." -ForegroundColor Green
            }
        }

        if ($this.Data.Count -le 0) {
            Write-Warning "No data captured to create a report"
        }
        return $this.Data
    }
}

([UrlRep]::new($FilePath, $APIKEY)).Check()

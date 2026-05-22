[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$APIKEY,
    [Parameter(Mandatory)]
    [string]$FilePath
)

. $PSScriptRoot\..\root\IOC.ps1

class UrlRep {
    [String]$URLValidator = "^(https?|hxxps?|ftp):\/\/[^\s/$.?#].[^\s]*$"
    [string]$ScanEndpoint = "https://www.virustotal.com/api/v3/urls"

    [string]$FilePath
    [System.Collections.Generic.HashSet[String]]$ListOfUrls
    [System.Collections.Generic.List[PSCustomObject]]$Data
    [Hashtable]$Headers = @{}

    UrlRep([string]$Path, [string]$Key) {
        $this.FilePath = $Path

        $this.ListOfUrls = New-Object System.Collections.Generic.HashSet[String]
        $this.Data = New-Object System.Collections.Generic.List[PSCustomObject]
        $this.Headers.Add("accept", "application/json")
        $this.Headers.Add("x-apikey", $Key)
        $this.Headers.Add("content-type", "application/x-www-form-urlencoded")
    }

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
                $this.ListOfUrls.Add($IOC) | Out-Null
            }
        }
    }

    [void] HandleWebException([object]$ExceptionObject, [string]$Url) {
        $Exception = $ExceptionObject.Exception
        $ErrorMessage = $ExceptionObject.Exception.Message
        $Resp = $exception.Response

        if ($null -ne $Resp) {
            $StatusCode = [int]$Resp.StatusCode

            switch ($StatusCode) {
                429 {
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

    [System.Collections.Generic.HashSet[String]] GetSelfLinkEndpoints() {
        [System.Collections.Generic.HashSet[String]]$SelfEndpoints = New-Object System.Collections.Generic.HashSet[String]        
        foreach ($Url in $this.ListOfUrls) {
            $SanitizedUrl = $Url.Replace(":", "[:]").Replace(".", "[.]")
            Write-Host "Url: $SanitizedUrl" -ForegroundColor Green
            
            $ContentType = "application/x-www-form-urlencoded"
            $EncodedUrl = "url=$([System.Uri]::EscapeDataString($Url))"
            try {
                $Response = Invoke-RestMethod -Method Post -Uri $this.ScanEndpoint -Headers $this.Headers -Body $EncodedUrl -ContentType $ContentType
                if ($Response.data) {
                    $D = $Response.data
                    $SelfEndpoints.Add($D.links.self) | Out-Null
                }
            }
            catch {
                $this.HandleWebException($_, $Url)
            }

        }
        return $SelfEndpoints
    }

    [System.Collections.Generic.List[PSCustomObject]] Check() {
        $this.ExtractUrls()
        Write-Host "$($this.ListOfUrls.Count) - Url(s) found in the file $($this.FilePath)"
        if ($this.ListOfUrls.Count -eq 0) { return $this.Data }
        Write-Host "Checking Url reputation..." -ForegroundColor Green

        [System.Collections.Generic.HashSet[String]]$SelfEndpoints = $this.GetSelfLinkEndpoints()
        foreach ($Endpoint in $SelfEndpoints) {
            try {
                $Response = Invoke-RestMethod -Method Get -Uri $Endpoint -Headers $this.Headers
                $D = $Response.data

                if ($D.attributes.status -ne "completed") {
                    Write-Host "`n`n`nScanning Status: $($D.attributes.status)" -ForegroundColor Yellow
                    Write-Host "Endpoint: $Endpoint" -ForegroundColor Yellow
                    Write-Host "Waiting for the analysis to complete..." -ForegroundColor Yellow
                    Write-Host "VirusTotal can take up to 2 minutes to analyze a Url, so we will retry every 10 seconds until we get a completed status or we reach the max retries" -ForegroundColor Yellow
                }

                $RetryCount = 1
                $MaxRetries = 12
                $IsReady = ($D.attributes.status -eq "completed")

                while (-not $IsReady -and $RetryCount -le $MaxRetries) {
                    Write-Host "Analysis pending... (Attempt: $RetryCount/$MaxRetries)" -ForegroundColor Yellow
                    Start-Sleep -Seconds 10
                    $Response = Invoke-RestMethod -Method Get -Uri $Endpoint -Headers $this.Headers
                    $D = $Response.data
                    
                    if ($D.attributes.status -eq "completed") {
                        Write-Host "Analysis completed!" -ForegroundColor Green
                        $IsReady = $true
                    }
                    $RetryCount++
                }

                if (-not $IsReady) {
                    Write-Warning "Unable to get the report for the endpoint: $Endpoint after 120 seconds."
                    continue 
                }

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
            catch {
                $this.HandleWebException($_, $Endpoint)
            }
        }

        if ($this.Data.Count -le 0) {
            Write-Warning "No data captured to create a report"
        }
        return $this.Data
    }
}

([UrlRep]::new($FilePath, $APIKEY)).Check()

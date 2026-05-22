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

    DomainRep([string]$Path, [string]$Key) {
        $this.FilePath = $Path

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
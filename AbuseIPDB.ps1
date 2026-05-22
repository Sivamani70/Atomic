[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$APIKEY,
    [Parameter(Mandatory)]
    [string]$FilePath
)

. $PSScriptRoot\root\IOC.ps1
. $PSScriptRoot\root\ExcelReport.ps1

class AbuseIPRep {
    [string]$IPV4Validator = "^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
    [String]$IPV6Validator = "^(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))$"
    [string]$ENDPOINT = "https://api.abuseipdb.com/api/v2/check"
    [string]$Type = "IP"

    [string]$FilePath
    [System.Collections.Generic.HashSet[String]]$ListOfIP
    [System.Collections.Generic.List[PSCustomObject]]$Data
    [Hashtable]$Headers = @{}

    AbuseIPRep([string]$Path, [string]$Key) {
        $this.FilePath = $Path

        $this.ListOfIP = New-Object System.Collections.Generic.HashSet[String]
        $this.Data = New-Object System.Collections.Generic.List[PSCustomObject]
        $this.Headers.Add("accept", "application/json")
        $this.Headers.Add("key", $Key)
    }

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
                $this.ListOfIP.Add($IOC) | Out-Null
            }
        }
    }

    [void] Check() {
        $this.ExtractIP()
        Write-Host "$($this.ListOfIP.Count) - IP(s) found in the file $($this.FilePath)"
        if ($this.ListOfIP.Count -eq 0) { return }
        Write-Host "Checking IP reputation..." -ForegroundColor Green

        foreach ($Ip in $this.ListOfIP) {
            Write-Host "IP: $Ip" -ForegroundColor Green

            [Hashtable]$QueryParameters = @{
                "ipAddress" = $Ip
                "verbose"   = " "
            }

            try {
                $Response = Invoke-RestMethod -Method Get -Uri $this.ENDPOINT -Headers $this.Headers -Body $QueryParameters
                $D = $Response.data
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

                    switch ($StatusCode) {
                        429 {
                            Write-Warning "VirusTotal API: Rate limit exceeded (429). Please wait before retrying."
                            break
                        }
                        Default {
                            Write-Warning "VirusTotal API: Received error code $StatusCode"
                            Write-Error "VirusTotal API: Received error Message $ErrorMessage while handling $Ip"
                        }
                    }
                }
            }
        }

        if ($this.Data.Count -le 0) {
            Write-Warning "No data captured to create a report"
            return
        }
        New-Report -IOCType $this.Type -Data $this.Data
    }
}

([AbuseIPRep]::new($FilePath, $APIKEY)).Check()
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$APIKEY,
    [Parameter(Mandatory)]
    [string]$FilePath
)

. $PSScriptRoot\..\root\IOC.ps1

class HashRep {
    [String]$MD5Validator = "^[a-fA-F0-9]{32}$"
    [String]$SHA1Validator = "^[a-fA-F0-9]{40}$"
    [String]$SHA256Validator = "^[a-fA-F0-9]{64}$"
    [String]$ENDPOINT = "https://www.virustotal.com/api/v3/files/"

    [string]$FilePath
    [System.Collections.Generic.HashSet[String]]$ListOfHashes
    [System.Collections.Generic.List[PSCustomObject]]$Data
    [Hashtable]$Headers = @{}

    HashRep([string]$Path, [string]$Key) {
        $this.FilePath = $Path

        $this.ListOfHashes = New-Object System.Collections.Generic.HashSet[String]
        $this.Data = New-Object System.Collections.Generic.List[PSCustomObject]
        $this.Headers.Add("accept", "application/json")
        $this.Headers.Add("x-apikey", $Key)
    }


    [void] ExtractHashes() {
        if (-not (Test-Path -Path $this.FilePath)) {
            Write-Error "File $($this.FilePath) is Invalid/File Not exist"
            return

        }
        Write-Host "Extracting Hashes"
        [System.Collections.Generic.HashSet[String]]$IOCs = Get-IOC_TXT -FilePath $this.FilePath
        foreach ($IOC in $IOCs) {
            $IOC = $IOC.Trim()
            if (($IOC -match $this.MD5Validator) -or ($IOC -match $this.SHA1Validator) -or ($IOC -match $this.SHA256Validator)) {
                $this.ListOfHashes.Add($IOC) | Out-Null
            }
        }
    }

    [System.Collections.Generic.List[PSCustomObject]] Check() {
        $this.ExtractHashes()
        Write-Host "$($this.ListOfHashes.Count) - Hash(es) found in the file $($this.FilePath)"
        if ($this.ListOfHashes.Count -eq 0) { return $this.Data }
        Write-Host "Checking Hash reputation..." -ForegroundColor Green

        foreach ($Hash in $this.ListOfHashes) {
            Write-Host "Hash: $Hash" -ForegroundColor Green
            [String]$FinalURL = $this.ENDPOINT + $Hash
            try {
                $Response = Invoke-RestMethod -Method Get -Uri $FinalURL -Headers $this.Headers
                $D = $Response.data
                [string]$BestName = "Unknown"              
                [string]$Description = "NA"                

                if ($D.attributes.signature_info) {
                    $Description = $D.attributes.signature_info.description                
                }

                if (-not [string]::IsNullOrWhiteSpace($D.attributes.meaningful_name)) {
                    $BestName = $D.attributes.meaningful_name
                }
                elseif ($null -ne $D.attributes.signature_info -and $D.attributes.signature_info["original name"]) {
                    $BestName = $D.attributes.signature_info["original name"]
                }
                elseif ($D.attributes.names.Count -gt 0) {
                    $BestName = $D.attributes.names[0]
                }

                $MD5 = $D.attributes.md5                
                $SHA1 = $D.attributes.sha1                
                $SHA256 = $D.attributes.sha256                
                $FileTag = $D.attributes.type_tag                
                $TypeDescription = $D.attributes.type_description                
                $Harmless = $D.attributes.last_analysis_stats.harmless
                $Malicious = $D.attributes.last_analysis_stats.malicious
                $Suspicious = $D.attributes.last_analysis_stats.suspicious
                $Undetected = $D.attributes.last_analysis_stats.undetected
                [int]$Total = $Harmless + $Malicious + $Suspicious + $Undetected
                [int]$MaliciousCount = $Malicious + $Suspicious

                $this.Data.Add([PSCustomObject]@{
                        FileName        = $BestName
                        Description     = $Description
                        MD5             = $MD5
                        SHA1            = $SHA1
                        SHA256          = $SHA256
                        FileTag         = $FileTag
                        TypeDescription = $TypeDescription
                        Result          = "$MaliciousCount//$Total"
                        TotalChecked    = $Total
                        Harmless        = $Harmless
                        Malicious       = $Malicious
                        Suspicious      = $Suspicious
                        Undetected      = $Undetected
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
                            Write-Error "VirusTotal API: Received error Message $ErrorMessage while handling $Hash"
                        }
                    }
                }
            }
        }

        if ($this.Data.Count -le 0) {
            Write-Warning "No data captured to create a report"
        }
        return $this.Data
    }    
}


Write-Host "Key: $APIKEY" -ForegroundColor Green
([HashRep]::new($FilePath, $APIKEY)).Check()

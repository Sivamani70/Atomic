<#
.SYNOPSIS
    Backend worker script for retrieving VirusTotal reputation scores for Hashes.

.DESCRIPTION
    This script defines and executes the `HashRep` class, which interacts with the 
    VirusTotal API v3. It extracts valid Hashes from a specified file, and queries their reputation. 

    If the Hash is not found in VirusTotal's database, it will log a warning with 404. The script also handles API rate limits gracefully.

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

class HashRep {
    [String]$MD5Validator = "^[a-fA-F0-9]{32}$"
    [String]$SHA1Validator = "^[a-fA-F0-9]{40}$"
    [String]$SHA256Validator = "^[a-fA-F0-9]{64}$"
    [String]$ENDPOINT = "https://www.virustotal.com/api/v3/files/"

    [string]$FilePath
    [System.Collections.Generic.HashSet[String]]$ListOfHashes
    [System.Collections.Generic.List[PSCustomObject]]$Data
    [Hashtable]$Headers = @{}
    [bool]$RateLimitHit


    HashRep([string]$Path, [string]$Key) {
        $this.FilePath = $Path
        $this.RateLimitHit = $false

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
            if ($this.RateLimitHit) {
                Write-Warning "VirusTotal API: Rate limit has been hit. Stopping further requests."
                break
            }

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
                        404 {
                            Write-Warning "VirusTotal API: Hash not found (404). The hash may not be recognized by VirusTotal."
                            break
                        }
                        429 {
                            $this.RateLimitHit = $true
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

([HashRep]::new($FilePath, $APIKEY)).Check()

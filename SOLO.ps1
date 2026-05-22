# SOLO - Separate Output Logging Orchestrator
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$APIKEY,
    [Parameter(Mandatory)]
    [string]$IOC_FilePath
)

. "$PSScriptRoot\root\ExcelReport.ps1"

function Get-IOCReport {

    $Scanners = @(
        @{ Label = "Hash"; Path = "virustotal\Hash.ps1" }
        @{ Label = "Domain"; Path = "virustotal\Domains.ps1" }
        @{ Label = "IP"; Path = "virustotal\IP.ps1" }
        @{ Label = "Url"; Path = "virustotal\Url.ps1" }
    )

    foreach ($Scanner in $Scanners) {
        $ScriptPath = Join-Path $PSScriptRoot $Scanner.Path
        [System.Collections.Generic.List[PSCustomObject]]$Data = . $ScriptPath -APIKEY $APIKEY -FilePath $IOC_FilePath
        if ($null -ne $Data -and $Data.Count -gt 0) {
            New-Report -IOCType $Scanner.Label -Data $Data
        }
    }
}

Get-IOCReport
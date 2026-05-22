# MIXR - Malware Investigation and eXcel Report
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$APIKEY,
    [Parameter(Mandatory)]
    [string]$IOC_FilePath
)

. "$PSScriptRoot\root\ReportSheet.ps1"
function Get-IOCReport {
    [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[PSCustomObject]]]$Response = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[PSCustomObject]]]::new()    

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
            $Response.Add($Scanner.Label, $Data)
        }
    }

    New-ExcelReport -Data $Response
}

if ((Test-Path -Path HKLM:\SOFTWARE\Microsoft\Office\*\Excel\)) {
    Get-IOCReport
}
else {
    Write-Host "Microsoft Excel is not installed on this system. Please install Excel to use this script." -ForegroundColor Red
}



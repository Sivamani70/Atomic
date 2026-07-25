<#
.SYNOPSIS
    Orchestrates malware indicator scans and generates isolated reports for each vector.
.DESCRIPTION
    SOLO (Separate Output Logging Orchestrator) acts as a specialized workflow engine.
    Unlike scripts that bundle data into a centralized workbook, SOLO iterates through 
    individual sub-scanners (Hash, Domain, IP, and URL modules) and dynamically triggers 
    an independent report call (`New-Report`) for each specific threat type as soon as 
    its data stream is collected.
.PARAMETER APIKEY
    The authorization key/token required by the underlying VirusTotal script scanners.
.PARAMETER IOC_FilePath
    The absolute or relative file path pointing to the text file containing raw threat indicators.
.EXAMPLE
    .\SOLO.ps1 -APIKEY "VT_Crypto_Token_XYZ" -IOC_FilePath "C:\Threats\today_iocs.txt"
.EXAMPLE
    .\SOLO.ps1 -APIKEY "VT API Key" -IOC_FilePath "C:\Threats\TMC_IOCs.txt"
.NOTES
    Author: SivaMani70
    Date: May 2026
    File Structure: Expects relative scanner modules located inside a 'virustotal' subfolder.
#>

# SOLO - Separate Output Logging Orchestrator
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$APIKEY,
    [Parameter(Mandatory)]
    [string]$IOC_FilePath
)

# Import the core reporting layout infrastructure script, which provides the New-Report function used to create individual Excel (or) CSV reports for each indicator type
. "$PSScriptRoot\root\ExcelOrCSVReport.ps1"

# --- Function: Get-IOCReport ---
# Purpose: Core coordinator block that maps distinct indicator types to sub-scripts,
# maintains rigid type integrity across execution streams, and routes 
# the accumulated dataset over to the Excel/CSV reporter.
function Get-IOCReport {

    # Manifest array defining task metadata and relative execution paths for each scanning subsystem
    $Scanners = @(
        @{ Label = "Hash"; Path = "virustotal\Hash.ps1" }
        @{ Label = "Domain"; Path = "virustotal\Domains.ps1" }
        @{ Label = "IP"; Path = "virustotal\IP.ps1" }
        @{ Label = "Url"; Path = "virustotal\Url.ps1" }
    )

    # Sequentially iterate through and execute each scanner module, passing the necessary parameters and capturing their outputs into the pre-defined dictionary structure for later aggregation
    foreach ($Scanner in $Scanners) {
        # Construct the absolute script path dynamically based on script execution origin and the relative path defined in the manifest, ensuring that the script can be executed from any location without hardcoded paths
        $ScriptPath = Join-Path $PSScriptRoot $Scanner.Path

        # Enforce strict type constraints on the receiver to capture data arrays safely from the sub-script pipeline and ensure that the expected data structure is maintained, allowing for early detection of any deviations in the execution flow
        [System.Collections.Generic.List[PSCustomObject]]$Data = . $ScriptPath -APIKEY $APIKEY -FilePath $IOC_FilePath

        # Verify telemetry records exist inside the type-safe collection before updating the dictionary mapping, preventing null reference exceptions and ensuring that only valid datasets are included in the final report compilation
        if ($null -ne $Data -and $Data.Count -gt 0) {
            New-Report -IOCType $Scanner.Label -Data $Data
        }
    }
}

# --- Script Entry Execution Hook ---
Get-IOCReport
<#
.SYNOPSIS
    Orchestrates multi-vector malware indicator scans and aggregates results into Excel.
.DESCRIPTION
    MIXR (Malware Investigation and eXcel Report) acts as a central execution controller.
    It loops through standalone sub-scripts (Hash, Domain, IP, and URL scanners), feeds them 
    the necessary API credentials, and collects their outputs using strict .NET generic 
    type constraints. Finally, it validates the presence of Microsoft Excel and compiles
    the structured data into a consolidated workbook.
.PARAMETER APIKEY
    The authorization key/token required by the underlying VirusTotal script scanners.
.PARAMETER IOC_FilePath
    The absolute or relative file path pointing to the text file containing raw threat indicators.
.EXAMPLE
    .\MIXR.ps1 -APIKEY "VT_Crypto_Token_XYZ" -IOC_FilePath "C:\Threats\today_iocs.txt"
.EXAMPLE
    .\MIXR.ps1 -APIKEY "VT API Key" -IOC_FilePath "C:\Threats\TMC_IOCs.txt"
.NOTES
    Author: SivaMani70
    Date: May 2026
    Prerequisites: Requires local installation of Microsoft Excel (validated via HKLM registry verification).
#>

# MIXR - Malware Investigation and eXcel Report
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$APIKEY,
    [Parameter(Mandatory)]
    [string]$IOC_FilePath
)

# Import the core Excel layout sheet generation infrastructure, which provides the New-ExcelReport function used to compile and format the final report
. "$PSScriptRoot\root\NewExcelReport.ps1"


# --- Function: Get-IOCReport ---
# Purpose: Core coordinator block that maps distinct indicator types to sub-scripts,
# maintains rigid type integrity across execution streams, and routes 
# the accumulated dataset over to the Excel reporter.
function Get-IOCReport {
    # Initialize a strictly-typed .NET Dictionary to bundle scanner labels with their respective data lists, ensuring that the data structure is consistent and type-safe for downstream processing
    [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[PSCustomObject]]]$Response = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[PSCustomObject]]]::new()    

    # Manifest array defining task metadata and relative execution paths for each scanning subsystem, 
    # enabling a modular and extensible architecture where new scanners can be added with minimal changes to the core logic
    $Scanners = @(
        @{ Label = "Hash"; Path = "virustotal\Hash.ps1" }
        @{ Label = "Domain"; Path = "virustotal\Domains.ps1" }
        @{ Label = "IP"; Path = "virustotal\IP.ps1" }
        @{ Label = "Url"; Path = "virustotal\Url.ps1" }
    )

    # Sequentially iterate through and execute each scanner module, passing the necessary parameters and capturing their outputs into the pre-defined dictionary structure for later aggregation
    foreach ($Scanner in $Scanners) {
        # Construct the absolute path dynamically based on script execution origin and the relative path defined in the manifest, ensuring that the script can be executed from any location without hardcoded paths
        $ScriptPath = Join-Path $PSScriptRoot $Scanner.Path

        # Enforce strict type constraints on the receiver to capture data arrays safely from the sub-script pipeline, ensuring that the expected data structure is maintained and that any deviations are caught early in the execution flow
        [System.Collections.Generic.List[PSCustomObject]]$Data = . $ScriptPath -APIKEY $APIKEY -FilePath $IOC_FilePath

        # Verify telemetry records exist inside the type-safe collection before updating the dictionary mapping, preventing null reference exceptions and ensuring that only valid datasets are included in the final report compilation
        if ($null -ne $Data -and $Data.Count -gt 0) {
            $Response.Add($Scanner.Label, $Data)
        }
    }

    # Pass the compiled multi-vector dataset payload to the workbook creation engine, which abstracts the complexities of Excel report generation and formatting, allowing for a clean separation of concerns between data collection and presentation logic
    New-ExcelReport -Data $Response
}

# --- Environment Lifecycle Verification ---
# Query registry hives to guarantee COM components exist before executing resource-intensive operations, ensuring that the script fails gracefully with user-friendly messages if prerequisites are not met
if ((Test-Path -Path HKLM:\SOFTWARE\Microsoft\Office\*\Excel\)) {
    Get-IOCReport
}
else {
    Write-Host "Microsoft Excel is not installed on this system. Please install Excel to use this script." -ForegroundColor Red
}



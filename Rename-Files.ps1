<#
.SYNOPSIS
    Renames Excel files in a target directory and formats them for Collector script ingestion.
.DESCRIPTION
    This script acts as a preprocessing utility. It validates a target directory path,
    prompts the user with a validated Yes/No loop, and sequentially renames files to a standardized 
    "A-N.xlsx" format. It then compiles the absolute file paths into a single comma-separated 
    string ready to copy directly into your Collector script configuration.
.PARAMETER Path
    The absolute or relative file path pointing to the folder containing your target files.
.EXAMPLE
    .\Rename-Files.ps1 -Path "D:\Projects\MalwareSamples"
    MalwareSamples may contain any number of excel files. This command will rename them to A-1.xlsx, A-2.xlsx, etc., and output a formatted string of their paths for Collector script use.
.NOTES
    Author: SivaMani70
    Date: May 2026
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [String] $Path
)


# --- Step 1: Environment & Directory Validation ---
if (-not (Test-Path -Path $Path)) {
    Write-Host "$Path is invalid/does not exist." -ForegroundColor Red
    return
}

Write-Host "This will rename all the files in the Folder: $Path" -ForegroundColor Yellow

# --- Step 2: Validated Input Loop via Do-Until ---
# This block enforces strict validation. The loop will refuse to break until the user
# provides an explicit 'y', 'yes', 'n', or 'no'.
do {
    $Response = Read-Host "Choose any option to proceed: (Y)es/(N)o"
    $Response = $Response.Trim().ToLower()
    
    if ([string]::IsNullOrEmpty($Response)) {
        Write-Host "Input cannot be empty. Please enter Y or N." -ForegroundColor Red
    }
} until ($Response -eq 'y' -or $Response -eq 'yes' -or $Response -eq 'n' -or $Response -eq 'no')

# --- Step 3: Sequential File Renaming Sequence ---
if ($Response[0] -eq 'y') {
    Write-Host "`nRenaming files..." -ForegroundColor Green
    $N = 1
    Get-ChildItem -Path $Path -Filter *.xlsx | ForEach-Object { 
        $OldName = $_.FullName
        
        # Enforce structural "A-N.xlsx" sequence name tagging
        Rename-Item -NewName "A-$N.xlsx" -Path $OldName
        $N++ 
    }
}
else {
    Write-Host "`nSkipping file renaming process as requested." -ForegroundColor Cyan
}

# --- Step 4: String Aggregation for TMC Ingestion ---
$Files = "" 
Get-ChildItem -Path $Path -Filter *.xlsx | ForEach-Object { 
    $Name = $_.FullName
    $Files = $Files + "'$Name'," 
}

# --- Step 5: Conditional Output Execution ---
# Automatically clean up the trailing comma if data exists
if ($Files.Length -gt 0) {
    $Files = $Files.TrimEnd(',')

    # Output the formatted string block for your TMC script
    Write-Host "`nUse the below line with TMC Script:`n" -ForegroundColor Green
    Write-Host $Files -ForegroundColor Cyan
}
else {
    # Fallback alert if zero .xlsx items were picked up by the provider
    Write-Host "`nNo Excel files to process with TMC." -ForegroundColor DarkCyan
}
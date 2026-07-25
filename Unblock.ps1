<#
.SYNOPSIS
    Recursively finds and unblocks all PowerShell script (.ps1) files in a target directory.

.DESCRIPTION
    This script searches through the specified directory path (and all subdirectories)
    for files ending with the '.ps1' extension. 
    Only run this script on files that you trust, as unblocking scripts can pose security risks if the scripts are malicious.

.PARAMETER Path
    Specifies the path to the directory containing the PowerShell scripts. 
    Defaults to the current working directory ('.').

.EXAMPLE
    .\Unblock.ps1

    Unblocks all .ps1 files in the current directory and its subdirectories.

.OUTPUTS
    Console messages showing the file path of each unblocked script.

.NOTES
    Requires execution permissions to run Unblock-File on downloaded or transferred files.
    Author: SivaMani70
    Date: July 2026
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Path = '.'
)

# Recursively locate all .ps1 files under the target path
Get-ChildItem -Path $Path -Filter "*.ps1" -Recurse | ForEach-Object {
    $FilePath = $_.FullName

    # Display progress message in the console
    Write-Host "Unblocking file: $FilePath" -ForegroundColor Green
    Unblock-File -Path $FilePath
}
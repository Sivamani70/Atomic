[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [String] $Path
)

if (!(Test-Path -Path $Path)) {
    Write-Host "$Path is invalid"
    return
}

Write-Host "This will rename all the files in the Folder: $Path" -ForegroundColor Yellow
$Response = Read-Host "Choose any option to proceed: (Y)es/(N)o"
$Response = $Response.ToLower();
if ($Response[0] -eq 'y') {
    Write-Host "Renaming files" -ForegroundColor Green
    $N = 1;
    Get-ChildItem -Path $Path | ForEach-Object { 
        $OldName = $_.FullName; 
        Rename-Item -NewName "A-$n.xlsx" -Path $OldName; 
        $N++; 
    }
}

$Files = ""; 
Get-ChildItem -Path $Path | ForEach-Object { 
    $Name = $_.FullName; 
    $Files = $Files + "'$Name',"; 
}

Write-Host "Use the below line with TMC Script (Remove the additional , at the end)`n`n$Files" -ForegroundColor Green

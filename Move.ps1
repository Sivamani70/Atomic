<#
.SYNOPSIS
    The main entry point to organize files in the current working directory. Organizes Excel and Text files into separate folders.
.DESCRIPTION
    Move-Files serves as the primary execution hook for this script module. It scans 
    the active current workspace directory (.), isolates all Microsoft Excel (*.xlsx) and 
    plain text (*.txt) files, ensures their respective destination folders exist, and 
    safely transfers them using a collision-aware renaming mechanism.

    A helper function called exclusively during the Move-Files lifecycle. If a file collision 
    is detected at the destination, it dynamically isolates the file components and appends a 
    "yyyyMMdd_HHmmss" execution timestamp string to make the transfer safe without overwriting data.
.EXAMPLE
    Import the file using dot sourcing and then execute the main function to sort files in the current directory.
    . .\Move.ps1
    Move-Files
    Run this command directly in your console to automatically sort your local files. (after the dot sourcing import)

.EXAMPLE
    . .\Move.ps1
    Move-Files
.LINK
    New-DirectoryIfMissing
    Move-ItemWithRename
#>

# Separating the function from the [script documentation] to allow help for the help command in powershell to work properly. 
# Get-Help .\Move.ps1 -Detailed
# Move.ps1 - File Organization and Management Script
function Move-Files {
    Write-Host "`n`nMoving (Excel and Text) files to respective folders..." -ForegroundColor Green
    $ExcelFiles = Get-ChildItem -Path . -Filter *.xlsx
    $ExcelFolder = "Excel Files"
    New-DirectoryIfMissing -Path $ExcelFolder
    foreach ($File in $ExcelFiles) {
        Move-ItemWithRename -File $File -DestinationFolder $ExcelFolder
    }

    $TextFiles = Get-ChildItem -Path . -Filter *.txt    
    $TextFolder = "Text Files"
    New-DirectoryIfMissing -Path $TextFolder
    foreach ($File in $TextFiles) {
        Move-ItemWithRename -File $File -DestinationFolder $TextFolder
    }
}
function Move-ItemWithRename($File, $DestinationFolder) {
    $Destination = Join-Path -Path $DestinationFolder -ChildPath $File.Name
    if (Test-Path -Path $Destination) {
        Write-Warning "File already exists: $Destination"
        $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
        $Extension = [System.IO.Path]::GetExtension($File.Name)
        $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $NewName = "$BaseName`_$Timestamp$Extension"
        $Destination = Join-Path -Path $DestinationFolder -ChildPath $NewName
    }
    Move-Item -Path $File.FullName -Destination $Destination
}

function New-DirectoryIfMissing($Path) {
    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Path $Path
    }
}




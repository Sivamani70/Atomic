
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


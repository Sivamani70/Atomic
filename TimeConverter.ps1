# This PowerShell script provides an interactive time zone conversion utility using Windows time zone IDs and the .NET TimeZoneInfo class.
# It allows users to convert a given date/time between multiple supported time zones via a menu-driven interface.
# EST and MST automatically adjusts for Daylight Saving Time (EDT) when applicable.
enum Zones {
    IST = 1
    UTC = 2
    MST = 3
    EST = 4
}

class Convert {
    hidden [string] $IST = "India Standard Time"
    hidden [string] $UTC = "UTC"
    hidden [string] $MST = "Mountain Standard Time"
    hidden [string] $EST = "Eastern Standard Time"

    Convert() {
        Write-Host -Object "Available Time Zones to Convert" -ForegroundColor "Green"
        Write-Host -Object "1. $($this.IST)" -ForegroundColor "Green"
        Write-Host -Object "2. $($this.UTC)" -ForegroundColor "Green"
        Write-Host -Object "3. $($this.MST)" -ForegroundColor "Green"
        Write-Host -Object "4. $($this.EST)" -ForegroundColor "Green"
    }

    [string] ZoneValue([Zones]$Value) {
        switch ($Value) {
            ([Zones]::IST) { return $this.IST }
            ([Zones]::UTC) { return $this.UTC }
            ([Zones]::MST) { return $this.MST }
            ([Zones]::EST) { return $this.EST }
            Default { return "None" }
        }
        return "None"
    }

    [void]Convert_Time([string]$From, [string]$To, [string]$GivenTime ) {
        $Source = [System.TimeZoneInfo]::FindSystemTimeZoneById($From)
        $Target = [System.TimeZoneInfo]::FindSystemTimeZoneById($To)
        $OriginalTime = Get-Date -Date $GivenTime
        # $OriginalTime = Get-Date -Date $GivenTime  -Format "dd-MM-yyyy hh:mm:ss tt"
        $ConvertedTime = [System.TimeZoneInfo]::ConvertTime($OriginalTime, $Source, $Target)
        $ConvertedTime = Get-Date -Date $ConvertedTime -Format 'dd-MM-yyyy hh:mm:ss tt'
        Write-Host -ForegroundColor "Green" -Object "Converted Time ($To):`t$ConvertedTime"
    }

    [PSCustomObject] Options() {
        [int]$From = Read-Host -Prompt "Choose the Input TimeZone [b/w 1 - 4]`t"
        [int]$To = Read-Host -Prompt "Choose the Output TimeZone [b/w 1 - 4]`t"
        [string]$DT = Read-Host -Prompt "Enter Date Time to convert`t[DD/MM/YYYY HH:MM:SS]"

        $FromZone = if ($From -le 4 -and $From -gt 0) { [Zones]$From } else { [Zones]::IST }
        $ToZone = if ($To -le 4 -and $To -gt 0) { [Zones]$To } else { [Zones]::UTC }

        write-Host -Object "Converting: $DT $FromZone -> $ToZone" -ForegroundColor "Green"

        return [PSCustomObject]@{
            From = $FromZone
            To   = $ToZone
            Time = $DT
        }
    }

    [void] Init() {
        $Obj = [PSCustomObject]($this.Options())
        $Source = $this.ZoneValue($Obj.From)
        $Destination = $this.ZoneValue($Obj.To)
        $this.Convert_Time($Source, $Destination, $Obj.Time)

    }

}

$convert = New-Object Convert
$convert.Init()
<#
.SYNOPSIS
    Provides an interactive utility to convert dates and times between supported time zones.

.DESCRIPTION
    This script utilizes a custom PowerShell class and the .NET [System.TimeZoneInfo] class 
    to facilitate time zone conversions. It presents an interactive menu allowing the user to 
    choose source and destination time zones (IST, UTC, MST, EST) and input a specific 
    date/time string. 
    
    Because it relies on native Windows Time Zone IDs, conversions involving regions with 
    Daylight Saving Time (like EST and MST) will automatically adjust to EDT/MDT when applicable 
    for the supplied date.

.EXAMPLE
    .\Convert-TimeZone.ps1
    
    Displays the menu and prompts the user for input:
    Available Time Zones to Convert
    1. India Standard Time
    2. UTC
    ...
    Choose the Input TimeZone [b/w 1 - 4]: 4
    Choose the Output TimeZone [b/w 1 - 4]: 1
    Enter Date Time to convert [DD/MM/YYYY HH:MM:SS]: 07/25/2026 14:00:00

.NOTES
    This script is fully interactive and relies on Read-Host for inputs. It does not accept 
    pipeline or command-line parameters.

    Author: SivaMani70
    Date: July 2026
#>




# This PowerShell script provides an interactive time zone conversion utility using Windows time zone IDs and the .NET TimeZoneInfo class.
# It allows users to convert a given date/time between multiple supported time zones via a menu-driven interface.
# EST and MST automatically adjusts for Daylight Saving Time (EDT) when applicable.

# Define an enumeration for the supported time zones to simplify menu selections
enum Zones {
    IST = 1
    UTC = 2
    MST = 3
    EST = 4
}

class Convert {
    # Hidden properties map the Enum values to exact Windows Time Zone IDs
    hidden [string] $IST = "India Standard Time"
    hidden [string] $UTC = "UTC"
    hidden [string] $MST = "Mountain Standard Time"
    hidden [string] $EST = "Eastern Standard Time"

    # Constructor: Displays the available zones whenever the class is instantiated
    Convert() {
        Write-Host -Object "Available Time Zones to Convert" -ForegroundColor "Green"
        Write-Host -Object "1. $($this.IST)" -ForegroundColor "Green"
        Write-Host -Object "2. $($this.UTC)" -ForegroundColor "Green"
        Write-Host -Object "3. $($this.MST)" -ForegroundColor "Green"
        Write-Host -Object "4. $($this.EST)" -ForegroundColor "Green"
    }

    # Maps the user's Enum selection to the corresponding Windows Time Zone ID string
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

    # Performs the actual conversion using .NET TimeZoneInfo
    [void]Convert_Time([string]$From, [string]$To, [string]$GivenTime ) {
        $Source = [System.TimeZoneInfo]::FindSystemTimeZoneById($From)
        $Target = [System.TimeZoneInfo]::FindSystemTimeZoneById($To)
        $OriginalTime = Get-Date -Date $GivenTime
        # $OriginalTime = Get-Date -Date $GivenTime  -Format "dd-MM-yyyy hh:mm:ss tt"
        $ConvertedTime = [System.TimeZoneInfo]::ConvertTime($OriginalTime, $Source, $Target)
        $ConvertedTime = Get-Date -Date $ConvertedTime -Format 'dd-MM-yyyy hh:mm:ss tt'
        Write-Host -ForegroundColor "Green" -Object "Converted Time ($To):`t$ConvertedTime"
    }

    # Interactively prompts the user for their conversion parameters
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

    # Orchestrator method: Gets options, translates IDs, and triggers the conversion
    [void] Init() {
        $Obj = [PSCustomObject]($this.Options())
        $Source = $this.ZoneValue($Obj.From)
        $Destination = $this.ZoneValue($Obj.To)
        $this.Convert_Time($Source, $Destination, $Obj.Time)

    }

}

# Instantiate the class and start the interactive prompt
$C = New-Object Convert
$C.Init()
enum Zones {
    IST = 1 
    UTC = 2 
    MST = 3 
}
 
class Convert {
    hidden [string] $IST = "India Standard Time" 
    hidden [string] $UTC = "UTC" 
    hidden [string] $MST = "Mountain Standard Time" 
 
    Convert() {
        Write-Host -Object "Available Time Zones to Convert" -ForegroundColor "Green" 
        Write-Host -Object "1. $($this.IST)" -ForegroundColor "Green" 
        Write-Host -Object "2. $($this.UTC)" -ForegroundColor "Green" 
        Write-Host -Object "3. $($this.MST)" -ForegroundColor "Green" 
    }
 
    [string] ZoneValue([Zones]$Value) {
        switch ($Value) { 
            ([Zones]::IST) { return $this.IST } 
            ([Zones]::UTC) { return $this.UTC } 
            ([Zones]::MST) { return $this.MST } 
            Default { return "None" } 
        }
        return "None"
    } 
    
 
    [void]Convert_Time([string]$From, [string]$To, [string]$GivenTime ) {
        $Source = [System.TimeZoneInfo]::FindSystemTimeZoneById($From) 
        $Target = [System.TimeZoneInfo]::FindSystemTimeZoneById($To) 
        $OriginalTime = Get-Date -Date $GivenTime 
        $ConvertedTime = [System.TimeZoneInfo]::ConvertTime($OriginalTime, $Source, $Target) 
        $ConvertedTime = Get-Date -Date $ConvertedTime -Format 'dd-MM-yyyy hh:mm:ss tt' 
        Write-Host -ForegroundColor "Green" -Object "Converted Time ($To):`t$ConvertedTime" 
    } 
 
 
 
    [PSCustomObject] Options() {
        [int]$From = Read-Host -Prompt "Choose the Input TimeZone [b/w 1 - 3]`t" 
        [int]$To = Read-Host -Prompt "Choose the Output TimeZone [b/w 1 - 3]`t" 
        [string]$DT = Read-Host -Prompt "Enter Date Time to convert`t[MM/DD/YYYY HH:MM:SS]" 
        
        $FromZone = if ($From -le 3 -and $From -gt 0) { [Zones]$From } else { [Zones]::IST }
        $ToZone = if ($To -le 3 -and $To -gt 0) { [Zones]$To } else { [Zones]::UTC }
 
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
 
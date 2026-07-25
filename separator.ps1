<#
.SYNOPSIS
    Parses, normalizes, categorizes, and deduplicates Indicators of Compromise (IOCs) from a text file.

.DESCRIPTION
    This script reads an input text file containing a raw list of IOCs and processes them 
    to create a clean, categorized output file (.\IOCs.txt). 
    
    During processing, the script automatically performs the following actions:
    - Reverses common sanitization techniques (e.g., replacing "[.]" with ".", "[:]" with ":").
    - Normalizes all entries by converting them to lowercase and trimming whitespace.
    - Deduplicates entries automatically by leveraging .NET HashSets.
    - Categorizes IOCs into distinct groups (MD5, SHA1, SHA256, Domains, URLs, Emails, IPs) using Regular Expressions.
    - Flags any unrecognized or un-parseable items for manual review at the bottom of the output file.

.PARAMETER Source
    Specifies the file path to the text file containing the raw, unformatted list of IOCs.

.EXAMPLE
    .\separator.ps1 -Source "C:\ThreatIntel\raw_ioc.txt"
    
    Reads 'raw_ioc.txt', processes the data, and outputs the categorized IOCs to '.\IOCs.txt' in the current directory.

.OUTPUTS
    Generates a text file named 'IOCs.txt' in the current working directory.

.NOTES
    Author: SivaMani70
    Date: May 2026
#>


[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [String]$Source
)


if (!(Test-Path -Path $Source)) {
    Write-Warning "$Source -- is not found in the present location"
    Write-Error "File Not Found"
    return
}

$Destination = ".\IOCs.txt"


[String]$MD5_Validator = "^[a-fA-F0-9]{32}$"
[String]$SHA1_Validator = "^[a-fA-F0-9]{40}$"
[String]$SHA256_Validator = "^[a-fA-F0-9]{64}$"
[String]$DomainValidator = "^(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})?$"   
[String]$URLValidator = "^(https?|hxxps?|ftp):\/\/[^\s/$.?#].[^\s]*$"
[String]$EmailValidator = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
[String]$IPV4Validator = "^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
[String]$IPV6Validator = "^(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))$"

[System.Collections.Generic.HashSet[String]]$ListOfIOCs = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]]$MD5 = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]]$SHA1 = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]]$SHA256 = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]]$Domains = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]]$URLS = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]]$IPS = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]]$Emails = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]]$OtherIOCs = New-Object System.Collections.Generic.HashSet[String]
[System.Text.StringBuilder]$Builder = New-Object -TypeName System.Text.StringBuilder

function Add-Items([System.Collections.Generic.HashSet[String]]$IOCs) {
    forEach ($IOC in $IOCs) {
        if ($IOC -ne "") {
            $Builder.Append($IOC + "`n") | Out-Null
        }
    }
}

function Write-File() {
    [string] $Data = ($Builder.ToString()).Trim()
    Set-Content -Path $Destination -Value $Data 
}

function Get-IOCData() {
    $ListOfIOCs = Get-Content -Path $Source
    forEach ($IOC in $ListOfIOCs) {
                
        $IOC = (([string]$IOC).ToLower()).Trim()
        
        if ($IOC -eq "") { Continue; }
        
        # Removing Sanitization 
        if ($IOC.Contains("[:]")) {
            $IOC = $IOC.Replace("[:]", ":")
        }

        if ($IOC.Contains("[.]")) {
            $IOC = $IOC.Replace("[.]", ".")
        }

        if ($IOC.Contains("[://]")) {
            $IOC = $IOC.Replace("[://]", "://")
        }

        #IP validation
        if ($IOC -match $IPV4Validator -or $IOC -match $IPV6Validator) {
            $IPS.Add($IOC) | Out-Null
            Continue;
        }

        #Domains validation
        if ($IOC -match $DomainValidator) {
            $Domains.Add($IOC) | Out-Null
            Continue;
        }

        #Emails validation
        if ($IOC -match $EmailValidator) {
            $Emails.Add($IOC) | Out-Null
            Continue;
        }

        #MD5 validation
        if ($IOC -match $MD5_Validator) {
            $MD5.Add($IOC) | Out-Null
            Continue;
        }

        #SHA1 validation
        if ($IOC -match $SHA1_Validator) {
            $SHA1.Add($IOC) | Out-Null
            Continue;
        }

        #SHA256 validation
        if ($IOC -match $SHA256_Validator) {
            $SHA256.Add($IOC) | Out-Null
            Continue;
        }

        #URL validation
        if ($IOC -match $URLValidator) {
            $URLS.Add($IOC) | Out-Null
            Continue;
        }
        $OtherIOCs.Add($IOC) | Out-Null
    }
}

# calling function
Get-IOCData


if ($MD5.Count -ne 0) {
    $Builder.Append("------------------MD5------------------" + "`n") | Out-Null
    Add-Items($MD5)
}
if ($SHA1.Count -ne 0) {
    $Builder.Append("------------------SHA1------------------" + "`n") | Out-Null
    Add-Items($SHA1)
}
if ($SHA256.Count -ne 0) {
    $Builder.Append("------------------SHA256------------------" + "`n") | Out-Null
    Add-Items($SHA256)
}
if ($Domains.Count -ne 0) {
    $Builder.Append("------------------Domains------------------" + "`n") | Out-Null
    Add-Items($Domains)
}
if ($URLS.Count -ne 0) {
    $Builder.Append("------------------URLS------------------" + "`n") | Out-Null
    Add-Items($URLS)
}
if ($Emails.Count -ne 0) {
    $Builder.Append("------------------Emails------------------" + "`n") | Out-Null
    Add-Items($Emails)
}
if ($IPS.Count -ne 0) {
    $Builder.Append("------------------IPS------------------" + "`n") | Out-Null
    Add-Items($IPS)
}

if ($OtherIOCs.Count -ne 0) {
    $Builder.Append(" " + "`n") | Out-Null
    $Builder.Append("------------------Review the Below IOCs Before Adding------------------" + "`n") | Out-Null
    Add-Items($OtherIOCs)
}

Write-File
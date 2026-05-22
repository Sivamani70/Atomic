[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [String] $Source
)


if (!(Test-Path -Path $Source)){
    Write-Warning "$Source -- is not found in the present location"
    Write-Error "File Not Found"
    return
}

$destination = ".\iocs.txt"


[String] $MD5_Validator = "^[a-fA-F0-9]{32}$"
[String] $SHA1_Validator = "^[a-fA-F0-9]{40}$"
[String] $SHA256_Validator = "^[a-fA-F0-9]{64}$"
[String] $DomainValidator = "^(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})?$"   
[String] $URLValidator = "^(https?|hxxps?|ftp):\/\/[^\s/$.?#].[^\s]*$"
[String] $EmailValidator = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
[String] $IPV4Validator = "^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
[String] $IPV6Validator = "^(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))$"

[System.Collections.Generic.HashSet[String]] $ListOfIOCs = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]] $MD5 = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]] $SHA1 = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]] $SHA256 = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]] $Domains = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]] $URLS = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]] $IPS = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]] $Emails = New-Object System.Collections.Generic.HashSet[String]
[System.Collections.Generic.HashSet[String]] $OtherIOCs = New-Object System.Collections.Generic.HashSet[String]
[System.Text.StringBuilder] $builder = New-Object -TypeName System.Text.StringBuilder

function Add-Items([System.Collections.Generic.HashSet[String]]$IOCs){
    forEach ($ioc in $IOCs){
        if ($ioc -ne ""){
            $builder.Append($ioc + "`n") | Out-Null
        }
    }
}

function WriteTo-File(){
    [string] $data = ($builder.ToString()).Trim()
    Set-Content -Path $destination -Value $data 
}

function Extract-IOCs(){
    $ListOfIOCs = Get-Content -Path $Source
    forEach ($ioc in $ListOfIOCs) {
                
        $ioc = ($ioc.ToLower()).Trim()
        
        if($ioc -eq ""){Continue;}
        
        # Removing Sanitization 
        if ($ioc.Contains("[:]")) {
            $ioc = $ioc.Replace("[:]", ":")
        }

        if ($ioc.Contains("[.]")) {
            $ioc = $ioc.Replace("[.]", ".")
        }

        if ($ioc.Contains("[://]")) {
            $ioc = $ioc.Replace("[://]", "://")
        }

        #IP validation
        if ($ioc -match $IPV4Validator -or $ioc -match $IPV6Validator) {
            $IPS.Add($ioc) | Out-Null
            Continue;
        }

        #Domains validation
        if ($ioc -match $DomainValidator) {
            $Domains.Add($ioc) | Out-Null
            Continue;
        }

        #Emails validation
        if ($ioc -match $EmailValidator) {
            $Emails.Add($ioc) | Out-Null
            Continue;
        }

        #MD5 validation
        if ($ioc -match $MD5_Validator) {
            $MD5.Add($ioc) | Out-Null
            Continue;
        }

        #SHA1 validation
        if ($ioc -match $SHA1_Validator) {
            $SHA1.Add($ioc) | Out-Null
            Continue;
        }

        #SHA256 validation
        if ($ioc -match $SHA256_Validator) {
            $SHA256.Add($ioc) | Out-Null
            Continue;
        }

        #URL validation
        if ($ioc -match $URLValidator) {
            $URLS.Add($ioc) | Out-Null
            Continue;
        }
        $OtherIOCs.Add($ioc) | Out-Null
    }
}

# calling function
Extract-IOCs


if ($MD5.Count -ne 0){
    $builder.Append("------------------MD5------------------" + "`n") | Out-Null
    Add-Items($MD5)
}
if ($SHA1.Count -ne 0){
    $builder.Append("------------------SHA1------------------" + "`n") | Out-Null
    Add-Items($SHA1)
}
if ($SHA256.Count -ne 0){
    $builder.Append("------------------SHA256------------------" + "`n") | Out-Null
    Add-Items($SHA256)
}
if ($Domains.Count -ne 0){
    $builder.Append("------------------Domains------------------" + "`n") | Out-Null
    Add-Items($Domains)
}
if ($URLS.Count -ne 0){
    $builder.Append("------------------URLS------------------" + "`n") | Out-Null
    Add-Items($URLS)
}
if ($Emails.Count -ne 0){
    $builder.Append("------------------Emails------------------" + "`n") | Out-Null
    Add-Items($Emails)
}
if ($IPS.Count -ne 0){
    $builder.Append("------------------IPS------------------" + "`n") | Out-Null
    Add-Items($IPS)
}

if($OtherIOCs.Count -ne 0){
    $builder.Append(" " + "`n") | Out-Null
    $builder.Append("------------------Review the Below IOCs Before Adding------------------" + "`n") | Out-Null
    Add-Items($OtherIOCs)
}

WriteTo-File
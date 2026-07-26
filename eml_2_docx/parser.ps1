param (
    [Parameter(Mandatory = $true)]
    [string]$EmlFilePath,
    [Parameter(Mandatory = $true)]
    [string]$IncidentNumber,
    [Parameter(Mandatory = $true)]
    [string]$OutFileDirectory
)


class NewDOC {

    # Entire doc font, size and colors are set here
    # If  you decided to change anything change the below values
    $WDColorAutomatic = -16777216
    $PurpleColor = 16744156
    $RED = 255
    $Font = "Open Sans" 
    $Heading = 18
    $Body = 11
    $SUBHeads = 14

    $FilePath = ""
    $IncidentNumber = ""
    $EmailData

    NewDOC([string]$Incident, [PSCustomObject]$Data, [string]$Name) {
        $this.EmailData = $Data

        # If there is no name for the incident is provided
        # $Incident variable is set to have the values of Current Time (Sec - Numbers only)
        if ($Incident.Trim() -eq "") {
            $Ticks = (Get-Date).Ticks
            Write-Warning "Provided Incident Number is empty"
            Write-Warning "Using a temporary PlaceHolder value {$Ticks}"
            $this.IncidentNumber = $Ticks
        }
        else {
            $this.IncidentNumber = $Incident
        }


        # If the directory provided does not exists
        # Then the output docx will be saved to current location
        # From where the script has been invoked
        if (-not(Test-Path $Name)) {
            $this.FilePath = Join-Path -Path (Get-Location).Path -ChildPath "$($this.IncidentNumber).docx"             
        }
        else {
            $this.FilePath = Join-Path -Path $Name -ChildPath "$($this.IncidentNumber).docx"  
        }
    }

    [Object[]] GetDomainReport([string]$Domain) {
        $ResolvedIPS = "IP Address`t`t:`t"
        $Reputation = "Reputation`t`t:`t"
        $CreationDate = "Creation Date`t`t:`t"
        $EXPDate = "Expire Date`t`t:`t"
        $AGE = "Domain Age`t`t:`t"
        $NameServers = "Name Servers`t`t:`t"

        # DNS Lookup using the PS cmdlet
        $Resolver = Resolve-DnsName -Name $Domain
        foreach ($r in $Resolver) {
            if ($r.Address) {
                $ResolvedIPS += "$(($r.Address).Trim()) | "
            }
            else {
                Write-Host "Unable to find the resolving IPs for $Domain" -ForegroundColor Red
                $ResolvedIPS = "IP Address`t`t:`tNA"
            }
        }
        $ResolvedIPS = $ResolvedIPS.TrimEnd("|", " ")

        
        # Domain LookUP in VirusTotal
        Write-Host "$Domain LookUP in VirusTotal" -ForegroundColor Green
        $Headers = @{}
        $Headers.Add("accept", "application/json")
        $Headers.Add("x-apikey", $env:VTAPIKEY)
        try {
            $VTresponse = Invoke-WebRequest -UseBasicParsing -Uri "https://www.virustotal.com/api/v3/domains/$Domain" -Method GET -Headers $Headers
            if ($VTresponse.StatusCode -eq 200) {
                $Data = $VTresponse.Content | ConvertFrom-Json
                if ($Data) {
                    $Stats = $Data.data.attributes.last_analysis_stats
                    $Mal = $Stats.malicious + $Stats.suspicious
                    $Total = $Stats.malicious + $Stats.suspicious + $Stats.undetected + $Stats.harmless + $Stats.timeout
                    $Reputation = "Reputation`t`t:`t$Mal/$Total"
                }
            }
            else {
                Write-Host "$Domain LookUP in VirusTotal: Failed" -ForegroundColor Red
                Write-Warning "Status Code: $($VTresponse.StatusCode)"
                Write-Warning "Response Data: $($VTresponse.Content)"
            }
        }
        catch [System.Net.WebException] {
            Write-Host "$Domain LookUP in VirusTotal: Failed" -ForegroundColor Red
            Write-Warning "Status Code $($_.Exception.Response.StatusCode)"
            Write-Warning $($_.Exception.Response)
        }
        catch {
            Write-Host "$Domain LookUP in VirusTotal: Failed" -ForegroundColor Red
            Write-Warning "Something went wrong"
        }



        try {
            # Domains WHOIS Lookup
            Write-Host "WHOIS LookUP for $Domain in IP2WHOIS" -ForegroundColor Green
            $Response = Invoke-WebRequest -UseBasicParsing -Method Get -Uri "https://api.ip2whois.com/v2?key=$($env:IP2LocationAPIKEY)&domain=$Domain"

            if ($Response.StatusCode -eq 200) {
                $Data = $Response.Content | ConvertFrom-Json
                if ($Data) {
                    $CreationDate = "Creation Date`t`t:`t$($Data.create_date)"
                    $EXPDate = "Expire Date`t`t:`t$($Data.expire_date)"
                    $AGE = "Domain Age`t`t:`t$([Math]::Round((($Data.domain_age) / 365), 2)) Years"
                    $NameServers = "Name Servers`t`t:`t$($Data.nameservers)"
                }
            }
            else {
                Write-Host "WHOIS LookUP for $Domain in IP2WHOIS: Failed" -ForegroundColor Red
                Write-Warning "Status Code: $($Response.StatusCode)"
                Write-Warning "Response Data: $($Response.Content)"
            }
        }
        catch [System.Net.WebException] {
            Write-Host "WHOIS LookUP for $Domain in IP2WHOIS: Failed" -ForegroundColor Red
            Write-Warning "Status Code $($_.Exception.Response.StatusCode)"
            Write-Warning $($_.Exception.Response)
        }
        catch {
            Write-Host "WHOIS LookUP for $Domain in IP2WHOIS: Failed" -ForegroundColor Red
            Write-Warning "Something went wrong"
        }

        # Domains Report
        return @(
            "Sender Domain`t:`t$Domain",
            $Reputation,
            $CreationDate,
            $EXPDate,
            $AGE,
            $NameServers
            $ResolvedIPS
        )
    }

    [void] Create() {

        $Domain = $this.EmailData['Domain'].ToLower()
        Write-Host "Sender Domain: `t$Domain"
        $DomainReport = $this.GetDomainReport($Domain)
        
        # Attachment Info
        $AttachmentCount = $this.EmailData['Attachment - Count']
        $AttachmentMalicious = "No"
        $AttachmentsAvailable = "No"
        if ($AttachmentCount -gt 0) { 
            $AttachmentMalicious = "No - [Check File Reputation]"
            $AttachmentsAvailable = "$AttachmentCount"
        }


        # Creating Word DOC
        $Word = New-Object -ComObject Word.Application
        $Word.Visible = $true
        $DOC = $Word.DOCuments.Add()

        # Change the Document font if you need at the constants sections
        $DOC.Content.Font.Name = $this.Font
        $Selection = $Word.Selection

        # Incident heading (centered, large, colored)
        $Selection.ParagraphFormat.Alignment = 1 # wdAlignParagraphCenter
        $Selection.Font.Size = $this.Heading
        $Selection.Font.Color = $this.PurpleColor 
        $Selection.TypeText("INCIDENT $($this.IncidentNumber)")
        $Selection.TypeParagraph()


        # Initial Analysis Intro
        $Selection.ParagraphFormat.Alignment = 0
        $Selection.Font.Size = $this.Body
        $Selection.Font.Color = $this.WDColorAutomatic 
        $Selection.TypeText("Please find the following initial analysis details.")
        $Selection.TypeParagraph()

        # Numbered List (details)
        # $details = @(
        #     "1. Subject     :`t$($this.EmailData['Subject'])",
        #     "2. Sender Id   :`t$($this.EmailData['From'])",
        #     "3. Recipient Id:`t$($this.EmailData['To'])",
        #     "4. Domain      :`t$($this.EmailData['Domain'])",
        #     "5. Blacklisted(Y/N):`t $($this.EmailData['Blacklisted'])",
        #     "6. Email Gateway:`t $($this.EmailData['EmailGateway'])",
        #     "7. Attachment(Y/N):`t $($this.EmailData['Attachment - Count'])",
        #     "8. Attachment (Malicious):`t --",
        #     "9. URL(Y/N):`t $($this.EmailData['URL'])",
        #     "10. URL(Malicious) (Y/N):`t  --",
        #     "11. Return Path:`t$($this.EmailData['Return-Path'])",
        #     "12. SCL:`t$($this.EmailData['SCL'])",
        #     "13. SPF Client IP:`t$($this.EmailData['SPF Client IP'])",
        #     "14. Reply-To:`t$($this.EmailData['Reply-To'])"
        # )

        #  foreach ($line in $details) {
        #     $Selection.Font.Size = $this.Body
        #     $Selection.Font.Color = $this.WDColorAutomatic
        #     $Selection.TypeText($line)
        #     $Selection.TypeParagraph()
        # }

        $details = @(
            # "Sent Time`t`t`t:`t$($this.EmailData['SentTime'])"
            "Received Time`t`t:`t$($this.EmailData['ReceivedTime'])"
            "Subject`t`t`t:`t$($this.EmailData['Subject'])",
            "Sender Id`t`t`t:`t$($this.EmailData['From'])",
            "Recipient Id`t`t`t:`t$($this.EmailData['To'])",
            "Domain`t`t`t:`t$Domain",
            "Blacklisted(Y/N)`t`t:`t$($this.EmailData['Blacklisted'])",
            "Email Gateway`t`t:`t$($this.EmailData['EmailGateway'])",
            "Attachment(Y/N)`t`t:`t$($this.EmailData['Attachment - Count'])",
            "Attachment (Malicious)`t:`t$AttachmentMalicious",
            "URL(Y/N)`t`t`t:`t$($this.EmailData['URL'])",
            "URL(Malicious) (Y/N)`t`t:`t--",
            "Return Path `t`t`t:`t$($this.EmailData['Return-Path'])",
            "SCL`t`t`t`t:`t$($this.EmailData['SCL'])",
            "SPF IP`t`t`t`t:`t$($this.EmailData['SPF Client IP'])",
            "Reply-To`t`t`t:`t$($this.EmailData['Reply-To'])"
        )

        $Selection.Range.ListFormat.ApplyNumberDefault()

        foreach ($line in $details) {
            $Selection.Font.Size = $this.Body
            $Selection.Font.Color = $this.WDColorAutomatic
            $Selection.TypeText($line)
            $Selection.TypeParagraph()
        }

        $Selection.Range.ListFormat.RemoveNumbers()

        # Blank line, "Ref:"
        $Selection.Font.Color = $this.PurpleColor
        $Selection.Font.Size = $this.Body
        $Selection.TypeText("Ref:")
        $Selection.TypeParagraph()
        # Body (placeholder)
        $Selection.Font.Bold = $false
        $Selection.TypeText("*")
        $Selection.TypeParagraph()

        # Domain Analysis Heading
        $Selection.Font.Bold = $true
        $Selection.Font.Size = $this.SUBHeads
        $Selection.TypeText("Domain Analysis")
        $Selection.TypeParagraph()
        # Domain Report
        $Selection.Font.Bold = $false
        foreach ($line in $DomainReport) {
            $Selection.Font.Size = $this.Body
            $Selection.Font.Color = $this.WDColorAutomatic
            $Selection.TypeText($line)
            $Selection.TypeParagraph()
        }

        # Section Header "Analysis"
        $Selection.Font.Color = $this.PurpleColor
        $Selection.Font.Bold = $true
        $Selection.Font.Size = $this.SUBHeads
        $Selection.TypeText("Analysis")
        $Selection.TypeParagraph()

        # Example analysis paragraph, with Words in red
        $Selection.Font.Color = $this.WDColorAutomatic
        $Selection.Font.Size = $this.Body
        $Selection.Font.Bold = $false
        $Selection.TypeText("User received a mail from")
        $Selection.Font.Color = $this.RED # red
        $Selection.TypeText(" $($this.EmailData['From']) ")
        $Selection.Font.Color = $this.WDColorAutomatic
        $Selection.TypeText("which was detected as a non-suspicious mail. As per the initial analysis we gathered that the mail came from ")
        $Selection.Font.Color = $this.RED # red
        $Selection.TypeText("$Domain")
        $Selection.Font.Color = $this.WDColorAutomatic
        $Selection.TypeText(".")
        $Selection.TypeParagraph()

        $Selection.TypeText("We also observed that there are *** URL(s) and $AttachmentsAvailable Attachment(s) in this email body.")
        $Selection.TypeParagraph()
        $Selection.TypeText("The Domain is clean as per virus total, Talos and URL void.")
        $Selection.TypeParagraph()

        # Ref: Link from URL void
        $URLVoid = "https://www.urlvoid.com/scan/$Domain/"
        $Selection.Font.Bold = $false
        $Selection.Font.Size = $this.Body
        $Selection.Font.Color = $this.RED
        $Selection.TypeText("Ref: ")
        $Selection.Font.Color = $this.WDColorAutomatic
        $Selection.TypeText($URLVoid)
        $Selection.TypeParagraph()

        # Ref: Link from virus total
        $VirusTotal = "https://www.virustotal.com/gui/domain/$Domain"
        $Selection.Font.Bold = $false
        $Selection.Font.Size = $this.Body
        $Selection.Font.Color = $this.RED
        $Selection.TypeText("Ref: ")
        $Selection.Font.Color = $this.WDColorAutomatic
        $Selection.TypeText($VirusTotal)
        $Selection.TypeParagraph()
        
        # Ref: Link from talos intelligence
        $Talos = "https://talosintelligence.com/reputation_center/lookup?search=$Domain"
        $Selection.Font.Bold = $false
        $Selection.Font.Size = $this.Body
        $Selection.Font.Color = $this.RED
        $Selection.TypeText("Ref: ")
        $Selection.Font.Color = $this.WDColorAutomatic
        $Selection.TypeText($Talos)
        $Selection.TypeParagraph()
        $Selection.TypeParagraph()
        
        # As per the analysis we observed that, there is ******** Attached file is an html document and is trying to get the credentials of the user. Intention of the mail is credential harvesting.
        # $Selection.Font.Bold = $false
        # $Selection.Font.Size = $this.Body
        # $Selection.Font.Color = $this.RED
        # $Selection.TypeText("As per the analysis we observed that, there is ******** Attached file is an html document and is trying to get the credentials of the user. Intention of the mail is credential harvesting.")
        # $Selection.TypeParagraph()
        # $Selection.TypeParagraph()

        # Security Team verdict
        # 	As per our Analysis, we have reached a verdict that the attached email is ***** Mail. 
        $Selection.Font.Bold = $false
        $Selection.Font.Size = $this.Body
        $Selection.Font.Color = $this.RED
        $Selection.TypeText("Security Team verdict:")
        $Selection.TypeParagraph()
        $Selection.Font.Color = $this.WDColorAutomatic
        $Selection.TypeText("`tAs per our Analysis, we have reached a verdict that the attached email is ")
        $Selection.Font.Color = $this.RED
        $Selection.TypeText("********")
        $Selection.Font.Color = $this.WDColorAutomatic
        $Selection.TypeText(" Mail.")

        # # Screenshots: 
        # $Selection.Font.Color = $this.PurpleColor
        # $Selection.Font.Size = $this.Body
        # $Selection.TypeText("Screenshots:")
        # $Selection.TypeParagraph()

        # Save and cleanup
        Write-Warning "Attempting to save the file at :$($this.FilePath)"
        $DOC.SaveAs($this.FilePath)
        $DOC.Close()
        $Word.Quit()
        Write-Warning "Closing word and releasing COM Objects"
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($DOC) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null

        Write-Host "Opening File: $($this.FilePath)"
        Start-Process -FilePath $this.FilePath

        $List = @($URLVoid, $VirusTotal, $Talos, "--new-window")
        # Start-Process -FilePath "msedge.exe" -ArgumentList $list
        Start-Process -FilePath "chrome.exe" -ArgumentList $list
    }
}

class EMLParser {

    [string]$File = ""
    # Only .eml and .msg are allowed
    [System.Object]$AllowedExtensions = @(".eml", ".msg")

    EMLParser([string]$EmlFilePath) {
        # Checking if the mail exists
        if (-not (Test-Path -Path $EmlFilePath)) {
            Write-Error "File Not Found $EmlFilePath"
            exit 1
        }

        # Checking if the file is valid eml or msg
        $Ext = [System.IO.Path]::GetExtension($EmlFilePath).ToLower()
        if (-not ($this.AllowedExtensions -contains $Ext)) {
            Write-Error "Error: File $EmlFilePath has unsupported extension $Ext. Exiting."
            exit 1
        }
        $this.File = $EmlFilePath
    }


    # Parsing the Mail to get the Information
    [PSCustomObject] Parser() {

        # Reading the mail contents as Raw to get some headers
        $EMLContent = Get-Content -Path $this.File -Raw
        $RawHeaders = $EMLContent -split "\r?\n\r?\n", 2 | Select-Object -First 1        
        $ReturnPath = ""
        $SCL = ""
        $SPF_IP = ""
        $ReplyTo = ""
        $AttachmentCount = 0

        # Looping through the raw headers to get ReturnPath, SPF_IP and SCL values
        foreach ($line in $RawHeaders -split "\r?\n") {
            if ($line -match "^Return-Path:\s*(.*)") {
                $ReturnPath = $Matches[1].Trim()
            }
            if ($line -match "^X-MS-Exchange-Organization-SCL:\s*(.*)") {
                $SCL = $Matches[1].Trim()
            }
            if ($line -match "^Authentication-Results:\s*(.*)") {
                $authResults += $Matches[1].Trim()
                if ($authResults -match "([\d\.]+)") {
                    $SPF_IP = $Matches[1]
                }
            }

            # if ($line -match "^DKIM-Signature:") {
            #     $dkim = "present"
            # }

        }

        # Creating Com Objects to read eml and msg files
        $Stream = New-Object -ComObject "ADODB.Stream"
        $Stream.Open()
        $Stream.LoadFromFile($this.File)

        $Email = New-Object -ComObject "CDO.Message"
        $Email.DataSource.OpenObject($Stream, "_Stream")

        # Getting the Return path from the stream object
        $RP = $Email.Fields.Item("urn:schemas:mailheader:return-path").Value
        # If failed to extract returnpath from the headers then get it from Stream objects
        if ($ReturnPath -eq "") { $ReturnPath = $RP }

        # $AuthResults = $Email.Fields.Item("urn:schemas:mailheader:authentication-results").Value
        # Write-Warning "Authentication-Results: $authResults"
        # Example Output
        # spf=pass (sender IP is 198.2.131.66) smtp.mailfrom=mail66.atl111.rsgsv.net; dkim=fail (body hash didnot verify) header.d=chaneyenterprises.com;dmarc=fail action=none header.from=ChaneyEnterprises.com;compauth=none reason=460
        # $AuthResults.Split(";") -> which will gives an array

        # Checking for attachments
        # Indexing starts from 1 -- Not sure why 😂
        $AttachmentCount = $Email.Attachments.Count
        if ($AttachmentCount -gt 0) {
            Write-Host "Number of attachments: $AttachmentCount"
    
            for ($i = 1; $i -le $AttachmentCount; $i++) {
                $attachment = $Email.Attachments.Item($i)
                Write-Warning "Attachment $($i) Name: $($attachment.FileName)"
                Write-Warning "Size: $($attachment.Size) bytes"
            }
        }
        else {
            Write-Host "No attachments found."
        }

        $ReplyTo = $Email.ReplyTo 
        if ($ReplyTo -eq "") {
            $ReplyTo = "NA"
        }

        $SentOn = $Email.SentOn
        $ReceivedTime = $Email.ReceivedTime
        $Subject = $Email.Subject
        # TODO: Modify these qualifiers [Based on requirment]
        # Only Below charecters are allowed in the Email        
        # '[^a-zA-Z0-9.\s@"()<>_\-]'
        $MailQualifiers = '[^a-zA-Z0-9.\s@"()<>_\-]'
        $From = $Email.From -replace $MailQualifiers
        $To = $Email.To
        # TODO: Modify these qualifiers [Based on requirment]
        # Only Below charecters are allowed in the Domain        
        # '[^a-zA-Z0-9._\-]'
        $DomainQualifiers = '[^a-zA-Z0-9._\-]'
        $SenderDomain = ($From.Split("@")[-1]) -replace $DomainQualifiers
        
        Write-Host "-------------------------------------"
        Write-Host "Sent ON:`t$SentOn"
        Write-Host "ReceivedTime:`t$ReceivedTime"
        Write-Host "Subject:`t$Subject"
        Write-Host "From:`t$From"
        Write-Host "To:`t$To"
        Write-Host "Reply-To:`t$ReplyTo"
        Write-Host "Return-Path:`t$ReturnPath"
        Write-Host "SCL:`t$SCL"
        Write-Host "SPF Client IP:`t$SPF_IP"

        # Releasing the COM Objects
        $Stream.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Stream) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Email) | Out-Null

        return @{
            SentTime             = "$SentOn"
            ReceivedTime         = "$ReceivedTime"
            Subject              = "$Subject"
            From                 = "$From"
            To                   = "$To"
            'Return-Path'        = "$ReturnPath"
            SCL                  = "$SCL"
            BCL                  = "--"
            SPF                  = "--"
            'SPF Client IP'      = "$SPF_IP"
            'Reply-To'           = "$ReplyTo"
            'Attachment - Count' = "$AttachmentCount"
            Blacklisted          = "No"
            EmailGateway         = "Delivered"
            URL                  = "--"
            Domain               = "$SenderDomain"
        }

    }
}



Write-Host "Import the API keys before running this using dot sourcing" -ForegroundColor Red
$EP = [EMLParser]::new($EmlFilePath) 
$Data = $EP.Parser()

$DOC = [NewDOC]::new($IncidentNumber, $Data, $OutFileDirectory)
$DOC.Create()
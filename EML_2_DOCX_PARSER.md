# Email to Word Document Parser (eml_2_docx)

This module provides enterprise-grade email parsing and incident report generation capabilities. It transforms raw EML (email) files into professionally formatted Word documents (.docx) with integrated threat intelligence analysis, including VirusTotal reputation checks and WHOIS domain information.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Module Structure](#module-structure)
- [Setup & Configuration](#setup--configuration)
- [Usage Guide](#usage-guide)
- [Output Report Format](#output-report-format)
- [API Requirements](#api-requirements)
- [Examples](#examples)
- [Report Elements](#report-elements)

---

## Overview

The `eml_2_docx` module automates the generation of incident response and email analysis reports. It:

- Parses email headers and content from EML files
- Extracts email metadata (sender, recipient, subject, attachments)
- Performs real-time threat intelligence lookups
- Generates professionally formatted Word documents for incident documentation
- Integrates DNS, VirusTotal, and WHOIS lookups automatically
- Enables rapid email incident triage and reporting

### Perfect For

- Email security incident response
- Phishing email analysis and documentation
- Malware/credential harvesting email triage
- Compliance and forensic reporting
- Email threat intelligence automation

---

## Features

### Core Capabilities

- **Email Parsing**: Extracts sender, recipient, subject, timestamps, and attachments info from EML files
- **Threat Intelligence Integration**:
  - **VirusTotal API**: Checks domain reputation and malicious/suspicious detection statistics
  - **IP2WHOIS API**: Retrieves domain WHOIS information (creation date, expiration, registrar, nameservers)
  - **DNS Resolution**: Performs standard DNS lookups for domain IP resolution
- **Professional Report Generation**: Creates formatted Word documents with:
  - Incident headers and numbering
  - Numbered email metadata sections
  - Domain analysis with intelligence data
  - Analysis narrative with highlighted findings
  - Security verdict section
  - Reference links to VirusTotal, URLVoid, and Talos Intelligence
- **Customizable Formatting**:
  - Configurable fonts, sizes, and colors
  - Professional purple headings and red highlights
  - Structured numbered lists and sections
  - Template-based document generation

### Report Elements

Each generated report includes:

| Element              | Description                                                                                                                                                   |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Incident Header**  | Large, centered heading with incident number                                                                                                                  |
| **Email Metadata**   | 14-item numbered list including: received time, subject, sender, recipient, domain, blacklist status, email gateway, attachments, URLs, SCL, SPF IP, reply-to |
| **Domain Analysis**  | DNS resolution results, VirusTotal reputation, WHOIS data (creation date, expiration, age, nameservers)                                                       |
| **Analysis Section** | Narrative analysis with highlighted sender domain and attachment counts                                                                                       |
| **Reference Links**  | Clickable references to URLVoid, VirusTotal, and Talos Intelligence for manual verification                                                                   |
| **Security Verdict** | Final security team assessment section                                                                                                                        |

---

## Module Structure

### Files Included

```
eml_2_docx/
├── parser.ps1          # Main email parser and report generator
└── keys.ps1            # API key configuration file
```

### Components

#### `parser.ps1`

The main PowerShell script containing:

- **`NewDOC` Class**:
  - Handles Word document generation and formatting
  - Manages email data extraction
  - Performs threat intelligence lookups
  - Generates professional formatted reports

- **Formatting Constants**:
  - `$Font`: Document font (default: "Open Sans")
  - `$Heading`: Main heading size (18pt)
  - `$SUBHeads`: Section heading size (14pt)
  - `$Body`: Body text size (11pt)
  - `$PurpleColor`: Section header color (16744156)
  - `$RED`: Highlight color for important text (255)

- **Key Methods**:
  - `GetDomainReport()`: Performs DNS, VirusTotal, and WHOIS lookups
  - `Create()`: Generates the final Word document

#### `keys.ps1`

Configuration file for API keys:

```powershell
$env:IP2LocationAPIKEY = "YOUR_API_KEY_HERE"
$env:VTAPIKEY = "YOUR_VIRUSTOTAL_API_KEY_HERE"
# $env:AbusseIPDBAPIKEY = "YOUR_ABUSEIPDB_API_KEY_HERE"  # Optional
```

---

## Setup & Configuration

### Prerequisites

- **PowerShell 5.0+**
- **Microsoft Word** (COM Interop required for .docx generation)
- **Internet Access** for API calls
- **API Keys** (see below)

### Step 1: Get API Keys

#### VirusTotal API Key

1. Visit [https://www.virustotal.com/](https://www.virustotal.com/)
2. Sign up or log in to your account
3. Navigate to your API key section
4. Copy your API key

#### IP2WHOIS API Key

1. Visit [https://www.ip2whois.com/](https://www.ip2whois.com/)
2. Create an account or log in
3. Access your API key from the dashboard
4. Copy your API key

### Step 2: Configure API Keys

Edit `keys.ps1` and replace the placeholder values:

```powershell
$env:IP2LocationAPIKEY = "your_ip2whois_api_key"
$env:VTAPIKEY = "your_virustotal_api_key"
```

Load the `keys.ps1` file once before executing `parser.ps1`. The `keys.ps1` script must be loaded once for each new PowerShell session.

```powershell
. .\keys.ps1
```

Alternatively, you can set environment variables directly before running: `Can omit the above step`

```powershell
$env:IP2LocationAPIKEY = "your_ip2whois_api_key"
$env:VTAPIKEY = "your_virustotal_api_key"
.\parser.ps1 -EmlFilePath "C:\Emails\suspicious.eml" -IncidentNumber "INC-2026-001" -OutFileDirectory "C:\Reports"
```

### Step 3: Prepare Email Files

Ensure you have EML format email files. Most email clients can export emails as EML:

- **Outlook**: Right-click email → Save As → "Outlook Message Format (\*.eml)"
- **Gmail**: Download as EML via forwarding to local mail client
- **Thunderbird**: Right-click email → Save As → Save as .eml

---

## Usage Guide

### Basic Syntax

```powershell
.\parser.ps1 -EmlFilePath <string> -IncidentNumber <string> -OutFileDirectory <string>
```

### Parameters

| Parameter          | Type   | Required | Description                                                                                                      |
| ------------------ | ------ | -------- | ---------------------------------------------------------------------------------------------------------------- |
| `EmlFilePath`      | string | Yes      | Full path to the EML file to parse and analyze                                                                   |
| `IncidentNumber`   | string | Yes      | Incident identifier (e.g., "INC-2026-001" or "PHO-20260725-001"). If empty, uses system ticks as placeholder     |
| `OutFileDirectory` | string | Yes      | Output directory for the generated Word document. If directory doesn't exist, saves to current working directory |

### Return Values

- **Success**: Generates a `.docx` file named `{IncidentNumber}.docx`
- **Validation**: If `IncidentNumber` is empty, generates a warning and uses timestamp
- **Location**: Saved to specified `OutFileDirectory` or current working directory

---

## Output Report Format

### Document Structure

```
┌─────────────────────────────────────┐
│   INCIDENT [INCIDENT_NUMBER]        │  (Purple, 18pt, centered)
│   (Heading)                         │
└─────────────────────────────────────┘

Please find the following initial analysis details.

1. Received Time           : [timestamp]
2. Subject                 : [email subject]
3. Sender Id               : [sender email]
4. Recipient Id            : [recipient email]
5. Domain                  : [sender domain]
6. Blacklisted(Y/N)        : [yes/no]
7. Email Gateway           : [gateway info]
8. Attachment(Y/N)         : [count]
9. Attachment (Malicious)  : [status]
10. URL(Y/N)               : [yes/no]
11. URL(Malicious) (Y/N)   : [status]
12. Return Path            : [return path]
13. SCL                    : [scam confidence level]
14. SPF IP                 : [ip address]
15. Reply-To               : [reply-to address]

Ref:
[Reference placeholder]

════════════════════════════════════
Domain Analysis
════════════════════════════════════

Sender Domain    : [domain]
Reputation       : [malicious+suspicious]/[total detections]
Creation Date    : [YYYY-MM-DD]
Expire Date      : [YYYY-MM-DD]
Domain Age       : [X.XX] Years
Name Servers     : [ns1.example.com, ns2.example.com]
IP Address       : [1.2.3.4, 5.6.7.8]

════════════════════════════════════
Analysis
════════════════════════════════════

User received a mail from [sender email] which was detected as
a non-suspicious mail. As per the initial analysis we gathered
that the mail came from [domain].

We also observed that there are *** URL(s) and [N] Attachment(s)
in this email body.

The Domain is clean as per virus total, Talos and URL void.

Ref: https://www.urlvoid.com/scan/[domain]/
Ref: https://www.virustotal.com/gui/domain/[domain]
Ref: https://talosintelligence.com/reputation_center/lookup?search=[domain]

════════════════════════════════════
Security Team verdict:
════════════════════════════════════

As per our Analysis, we have reached a verdict that the attached
email is ******** Mail.

[Analyst Notes - Manual addition required]
```

---

## API Requirements

### VirusTotal API

**Endpoint**: `https://www.virustotal.com/api/v3/domains/{domain}`

**What it provides**:

- Malicious/Suspicious detection count
- Total analysis count
- Security vendor consensus

**Rate Limits**: Depends on subscription tier

- Free tier: 4 requests/minute
- Premium tiers: Higher limits

**Error Handling**:

- Returns HTTP 200 on success
- Returns appropriate HTTP error codes on failure
- Script logs warnings for API failures and continues processing

### IP2WHOIS API

**Endpoint**: `https://api.ip2whois.com/v2?key={api_key}&domain={domain}`

**What it provides**:

- Domain creation date
- Domain expiration date
- Domain age calculation
- Nameserver information
- Registrar details

**Rate Limits**: Depends on subscription tier

- Free tier: 500 requests/month
- Paid tiers: Higher quotas

**Error Handling**:

- Returns HTTP 200 on success
- Script logs warnings on API failures
- Processing continues with "NA" values if WHOIS lookup fails

### DNS Resolution

**Method**: Windows native `Resolve-DnsName` PowerShell cmdlet

**What it provides**:

- IP address(es) for the domain
- No authentication required
- Local system DNS resolver used

---

## Examples

### Example 1: Basic Email Analysis

```powershell
# Set API keys
$env:VTAPIKEY = "your_vt_key"
$env:IP2LocationAPIKEY = "your_ip2whois_key"

# Parse a suspicious email
.\parser.ps1 -EmlFilePath "C:\SecurityIncidents\phishing_email.eml" `
             -IncidentNumber "PHO-2026-0725-001" `
             -OutFileDirectory "C:\SecurityReports\2026-07\Phishing"
```

**Output**: `C:\SecurityReports\2026-07\Phishing\PHO-2026-0725-001.docx`

### Example 2: Batch Processing Multiple Emails

```powershell
# Set API keys
$env:VTAPIKEY = "your_vt_key"
$env:IP2LocationAPIKEY = "your_ip2whois_key"

# Create output directory
$OutputDir = "C:\SecurityReports\Daily_Analysis"
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

# Process multiple emails
$EmlFiles = Get-ChildItem -Path "C:\EmailQueue" -Filter "*.eml"
$IncidentCounter = 1

foreach ($Email in $EmlFiles) {
    $IncidentNumber = "EMAIL-2026-{0:D4}" -f $IncidentCounter

    .\parser.ps1 -EmlFilePath $Email.FullName `
                 -IncidentNumber $IncidentNumber `
                 -OutFileDirectory $OutputDir

    $IncidentCounter++
}
```

### Example 3: Using Current Directory for Output

```powershell
# If output directory doesn't exist, file saves to current location
.\parser.ps1 -EmlFilePath ".\suspicious_email.eml" `
             -IncidentNumber "INC-001" `
             -OutFileDirectory "C:\NonExistentPath"

# Output: .\INC-001.docx (in current working directory)
```

### Example 4: Empty Incident Number (Not Recommended)

```powershell
# If incident number is empty, script generates timestamp-based ID
.\parser.ps1 -EmlFilePath ".\email.eml" `
             -IncidentNumber "" `
             -OutFileDirectory "C:\Reports"

# Output: Warning message + .\[TIMESTAMP].docx
# Example: .\637334052341234567.docx
```

---

## Report Elements

### Email Metadata Extraction

The parser extracts the following email properties:

| Property               | Source                             | Used In Report                                      |
| ---------------------- | ---------------------------------- | --------------------------------------------------- |
| **SentTime**           | Email header: Date                 | Metadata section (commented out in current version) |
| **ReceivedTime**       | Email header: Received             | Metadata item #1                                    |
| **Subject**            | Email header: Subject              | Metadata item #2                                    |
| **From**               | Email header: From                 | Metadata item #3, Analysis section                  |
| **To**                 | Email header: To                   | Metadata item #4                                    |
| **Domain**             | Extracted from sender email domain | Throughout report, for lookups                      |
| **Return-Path**        | Email header: Return-Path          | Metadata item #12                                   |
| **Reply-To**           | Email header: Reply-To             | Metadata item #15                                   |
| **Blacklisted**        | Custom field                       | Metadata item #6                                    |
| **EmailGateway**       | Custom field                       | Metadata item #7                                    |
| **Attachment - Count** | Count of email attachments         | Metadata item #8, Analysis section                  |
| **URL**                | URL presence in email              | Metadata item #10                                   |
| **SCL**                | Spam Confidence Level (Exchange)   | Metadata item #13                                   |
| **SPF Client IP**      | SPF authentication IP              | Metadata item #14                                   |

### Domain Analysis Data

For the sender's domain, the report includes:

| Data Point          | Source API            | Purpose                                     |
| ------------------- | --------------------- | ------------------------------------------- |
| **IP Address**      | DNS Resolution        | Network infrastructure verification         |
| **Reputation**      | VirusTotal            | Malicious/Suspicious detection count        |
| **Creation Date**   | IP2WHOIS              | Domain age and registration history         |
| **Expiration Date** | IP2WHOIS              | Domain lifecycle information                |
| **Domain Age**      | IP2WHOIS (calculated) | Risk indicator; younger domains higher risk |
| **Name Servers**    | IP2WHOIS              | Authoritative DNS information               |

### Threat Intelligence References

The report automatically includes reference links for:

1. **URLVoid**: `https://www.urlvoid.com/scan/{domain}/`
2. **VirusTotal**: `https://www.virustotal.com/gui/domain/{domain}`
3. **Talos Intelligence**: `https://talosintelligence.com/reputation_center/lookup?search={domain}`

These enable analysts to perform additional manual verification if needed.

---

## Customization

### Changing Document Formatting

To customize the appearance of generated reports, edit the formatting constants in `parser.ps1`:

```powershell
# In the NewDOC class:
$Font = "Calibri"              # Change document font
$Heading = 24                  # Change main heading size
$Body = 12                     # Change body text size
$SUBHeads = 16                 # Change section heading size
$PurpleColor = 8421504         # Change primary heading color
$RED = 255                     # Change highlight color (RGB)
```

### Common Color Codes

- Red: `255`
- Green: `32768`
- Blue: `16711680`
- Purple: `8421504` or `16744156`
- Black: `-16777216`

### Modifying Report Sections

To add or modify report sections, edit the `Create()` method in the `NewDOC` class. Key operations:

```powershell
$Selection.TypeText("Your text here")           # Add text
$Selection.TypeParagraph()                      # New line
$Selection.Font.Size = $this.Body              # Set font size
$Selection.Font.Color = $this.RED              # Set color
$Selection.Font.Bold = $true                   # Bold text
$Selection.ParagraphFormat.Alignment = 1       # Center (0=left, 1=center, 2=right)
```

---

## Troubleshooting

### Issue: "Cannot find Word application"

**Cause**: Microsoft Word is not installed or COM interop is unavailable

**Solution**:

- Install Microsoft Office with Word
- Ensure Word is not blocked by security policies
- Try running as Administrator

### Issue: "API Key Error" or HTTP 401

**Cause**: Invalid or expired API keys

**Solution**:

- Verify API keys in `keys.ps1`
- Check API key expiration dates
- Ensure correct API key is used (VT key vs IP2WHOIS key)
- Regenerate API keys if necessary

### Issue: "Domain not found in VirusTotal"

**Cause**: Domain hasn't been indexed by VirusTotal yet or doesn't have reputation data

**Solution**:

- This is normal for newly registered domains
- Script will display "NA" or generic values
- Continue with WHOIS and DNS data that was retrieved

### Issue: Output file not created in expected location

**Cause**: Output directory path doesn't exist or invalid permissions

**Solution**:

- Create output directory before running script
- Check folder permissions allow write access
- Use absolute paths instead of relative paths
- If directory doesn't exist, file saves to current working directory (by design)

### Issue: "Zombie" Word processes remain after completion

**Cause**: Word COM objects not properly released

**Solution**:

- The script includes COM cleanup, but if processes persist:
  - Manually close Word: `Get-Process WINWORD | Stop-Process -Force`
  - Or restart PowerShell session

---

## Performance Considerations

- **API Lookups**: Each email processing performs 2-3 API calls (VirusTotal + WHOIS). Plan for network latency.
- **Rate Limiting**: Be aware of API rate limits when batch processing
- **DNS Lookups**: Local DNS queries are fast but depend on network configuration
- **Word COM Automation**: Document generation is I/O intensive; allow 5-10 seconds per email

### Batch Processing Optimization

```powershell
# Process emails with delays to avoid rate limiting
$EmlFiles = Get-ChildItem -Path "C:\Emails" -Filter "*.eml"

foreach ($Email in $EmlFiles) {
    .\parser.ps1 -EmlFilePath $Email.FullName `
                 -IncidentNumber $Email.BaseName `
                 -OutFileDirectory "C:\Reports"

    Start-Sleep -Seconds 2  # 2-second delay between emails
}
```

---

## Integration with Other Tools

The `eml_2_docx` module complements other tools in the suite:

| Tool                  | Integration | Purpose                                                                |
| --------------------- | ----------- | ---------------------------------------------------------------------- |
| **SOLO/MIXR**         | Manual      | Extract IOCs (hashes, URLs, IPs) from email and analyze with SOLO/MIXR |
| **AbuseIPDB**         | Manual      | Use sender/relay IPs from email headers with AbuseIPDB.ps1             |
| **separator.ps1**     | Manual      | Categorize extracted IOCs from email body/attachments                  |
| **TimeConverter.ps1** | Manual      | Normalize email timestamps for timeline reconstruction                 |

---

## Author

Created by **SivaMani70** (May 2026)

## See Also

- [Main README](README.md) - Core IOC analysis tools
- [POWERSHELL_UTILITIES.md](POWERSHELL_UTILITIES.md) - Additional utility scripts
- [SOLO.ps1](SOLO.ps1) - Single-engine IOC analyzer
- [MIXR.ps1](MIXR.ps1) - Multi-engine IOC analyzer
- [AbuseIPDB.ps1](AbuseIPDB.ps1) - IP reputation analyzer

# VirusTotal IOC Analysis Suite

A PowerShell-based toolkit for analyzing Indicators of Compromise (IOCs) against VirusTotal's threat intelligence database. This project provides two distinct workflows for threat assessment: **SOLO** (for isolated reports) and **MIXR** (for consolidated Excel workbooks).

## 🎯 Overview

This suite streamlines malware investigation workflows by automating the scanning of multiple threat indicator types (Hashes, Domains, IPs, and URLs) against VirusTotal's API, with additional IP intelligence from AbuseIPDB. It intelligently routes findings into either independent reports or a unified Excel workbook, enabling security teams to rapidly triage and analyze security threats.

## 📋 Features

- **Multi-Vector Scanning**: Analyze hashes (MD5, SHA-1, SHA-256), domains, IP addresses, and URLs simultaneously
- **IP Intelligence Integration**: Query AbuseIPDB for comprehensive IP reputation and abuse reports
- **Three Scanning Options**:
  - **SOLO**: Generates separate reports for each indicator type as data becomes available
  - **MIXR**: Aggregates all data into a consolidated Excel workbook
  - **AbuseIPDB**: Dedicated IP reputation and abuse intelligence gathering
- **API Rate Limiting**: Handles VirusTotal and AbuseIPDB API throttling gracefully
- **Robust Error Handling**: Validates indicators and logs warnings for missing or invalid data
- **Modular Architecture**: Easy to extend with additional scanner modules
- **Type-Safe Processing**: Uses .NET generics for strict data integrity
- **Automatic De-duplication**: Uses .NET HashSet for efficient unique indicator handling

## 🚀 Getting Started

### Prerequisites

- **PowerShell 5.0+** (Windows PowerShell or PowerShell Core)
- **VirusTotal API Key** (Get one from [virustotal.com](https://www.virustotal.com/))
- **AbuseIPDB API Key** (Optional; Get one from [abuseipdb.com](https://www.abuseipdb.com/))
- **Microsoft Excel** (Required for MIXR; not required for SOLO or AbuseIPDB standalone)

### Installation

1. Clone or download the repository:
   ```powershell
   git clone <repository-url>
   cd readme-virustotal-solo-mixer
   ```

2. Prepare your IOC file (a plain text file with one indicator per line):
   ```
   4d967a2a964760d93477d2d023efb0b1
   example.com
   192.0.2.1
   https://suspicious-domain.com/malware.exe
   ```

### API Key Setup

Store your VirusTotal API key securely. You can either:
- Pass it directly as a parameter
- Store it in an environment variable and reference it

## 📖 Usage

### Option 1: SOLO (Separate Output Logging Orchestrator)

Use SOLO when you want **independent reports for each indicator type** as they complete scanning:

```powershell
.\SOLO.ps1 -APIKEY "your_virustotal_api_key" -IOC_FilePath "C:\Threats\today_iocs.txt"
```

**What SOLO does:**
- Scans through Hash, Domain, IP, and URL modules sequentially
- Generates a separate report immediately after each module completes
- Returns to the command line as soon as processing is done
- Ideal for CI/CD pipelines or when you need results as they arrive

**Example with environment variable:**
```powershell
$env:VT_API_KEY = "your_virustotal_api_key"
.\SOLO.ps1 -APIKEY $env:VT_API_KEY -IOC_FilePath "C:\Threats\daily_iocs.txt"
```

### Option 2: MIXR (Malware Investigation and eXcel Report)

Use MIXR when you want a **unified Excel workbook with all results**:

```powershell
.\MIXR.ps1 -APIKEY "your_virustotal_api_key" -IOC_FilePath "C:\Threats\today_iocs.txt"
```

**What MIXR does:**
- Scans all indicator types (Hash, Domain, IP, URL)
- Aggregates results into a structured data dictionary
- Generates a professionally formatted Excel workbook with multiple sheets (one per indicator type)
- Requires Microsoft Excel to be installed on the system

**Example with relative path:**
```powershell
.\MIXR.ps1 -APIKEY "your_virustotal_api_key" -IOC_FilePath ".\iocs\malware_hashes.txt"
```

### Option 3: AbuseIPDB (IP Reputation & Abuse Intelligence)

Use AbuseIPDB when you need **detailed IP reputation and abuse report data** from the AbuseIPDB database:

```powershell
.\AbuseIPDB.ps1 -APIKEY "your_abuseipdb_api_key" -FilePath "C:\Threats\suspicious_ips.txt"
```

**What AbuseIPDB does:**
- Extracts and de-duplicates IPv4 and IPv6 addresses from your IOC file
- Queries the AbuseIPDB v2 API for comprehensive reputation data
- Returns detailed abuse reports including:
  - Total abuse reports count
  - Abuse confidence score (0-100%)
  - ISP and domain information
  - Country code and name
  - Tor exit node detection
  - Whitelisting status
  - Usage type classification
- Generates a detailed Excel/CSV report with all findings

**Example with IP-only file:**
```powershell
.\AbuseIPDB.ps1 -APIKEY "your_abuseipdb_api_key" -FilePath "C:\Threats\ips_to_check.txt"
```

**Key Features of AbuseIPDB:**
- Automatic de-duplication of IPs (no duplicates in results)
- Validates both IPv4 and IPv6 formats rigorously
- Graceful handling of rate limits with user notifications
- Extracts only IP addresses from mixed IOC files

## 📊 Supported Indicator Types

| Indicator Type | Format | Example | Scanners |
|---|---|---|---|
| **Hash** | MD5, SHA-1, SHA-256 | `4d967a2a964760d93477d2d023efb0b1` | VirusTotal (SOLO, MIXR) |
| **Domain** | FQDN | `example.com`, `subdomain.example.com` | VirusTotal (SOLO, MIXR) |
| **IP Address** | IPv4 or IPv6 | `192.0.2.1`, `2001:db8::1` | VirusTotal (SOLO, MIXR), AbuseIPDB |
| **URL** | Full URI | `https://example.com/path/malware.exe` | VirusTotal (SOLO, MIXR) |

**Note**: AbuseIPDB specializes in IP reputation and provides different data than VirusTotal's IP scanner. Use AbuseIPDB for detailed abuse reports and historical threat data.

## 📁 Project Structure

```
readme-virustotal-solo-mixer/
├── SOLO.ps1                          # Main entry point for isolated reports
├── MIXR.ps1                          # Main entry point for Excel workbook
├── AbuseIPDB.ps1                     # IP reputation and abuse intelligence scanner
│
├── virustotal/                       # VirusTotal scanner modules
│   ├── Hash.ps1                      # Hash reputation scanner
│   ├── Domains.ps1                   # Domain reputation scanner
│   ├── IP.ps1                        # IP reputation scanner
│   └── Url.ps1                       # URL reputation scanner
│
├── root/                             # Core infrastructure
│   ├── IOC.ps1                       # IOC parsing and validation utilities
│   ├── NewExcelReport.ps1            # Excel workbook generation
│   ├── ExcelOrCSVReport.ps1          # Individual report generation (CSV/Excel)
│   └── Prompts.ps1                   # User prompt utilities
│
└── README.md                         # This file
```

## 🔄 How It Works

### SOLO Workflow

```
IOC File Input
    ↓
┌─────────────────────────────────────┐
│  Iterate Through Scanner Modules    │
└─────────────────────────────────────┘
    ↓
┌─────────────────────────────────────┐
│  1. Hash.ps1 → Generate Report      │
│  2. Domains.ps1 → Generate Report   │
│  3. IP.ps1 → Generate Report        │
│  4. Url.ps1 → Generate Report       │
└─────────────────────────────────────┘
    ↓
Individual Reports (CSV or Excel)
```

### MIXR Workflow

```
IOC File Input
    ↓
┌─────────────────────────────────────┐
│  Iterate Through Scanner Modules    │
└─────────────────────────────────────┘
    ↓
┌─────────────────────────────────────┐
│  Aggregate All Results Into         │
│  Dictionary Structure               │
└─────────────────────────────────────┘
    ↓
┌─────────────────────────────────────┐
│  Verify Excel is Installed          │
└─────────────────────────────────────┘
    ↓
Unified Excel Workbook with Multiple Sheets
```

### AbuseIPDB Workflow

```
IOC File Input (Mixed Indicators)
    ↓
┌─────────────────────────────────────┐
│  Extract & Validate IP Addresses    │
│  (IPv4 & IPv6 Format Checking)      │
└─────────────────────────────────────┘
    ↓
┌─────────────────────────────────────┐
│  De-Duplicate IPs Using HashSet     │
└─────────────────────────────────────┘
    ↓
┌─────────────────────────────────────┐
│  Query Each IP via AbuseIPDB v2 API │
│  (Handle Rate Limits & Errors)      │
└─────────────────────────────────────┘
    ↓
┌─────────────────────────────────────┐
│  Compile Reputation & Abuse Data    │
└─────────────────────────────────────┘
    ↓
Excel/CSV Report with IP Intelligence
```

## 🔑 Parameters

### Common Parameters

#### `-APIKEY` (Mandatory)
- **Type**: `string`
- **Description**: Your API authentication key (VirusTotal or AbuseIPDB depending on the script)
- **Example**: `"VT_Crypto_Token_XYZ"` or `"AbuseIPDB_Secret_Token_123"`

#### `-IOC_FilePath` / `-FilePath` (Mandatory)
- **Type**: `string`
- **Description**: Absolute or relative path to the text file containing indicators of compromise (one per line)
- **Example**: `"C:\Threats\today_iocs.txt"` or `".\iocs\malware.txt"`
- **Note**: Parameter name is `-IOC_FilePath` for SOLO/MIXR and `-FilePath` for AbuseIPDB

### Script-Specific Parameters

| Script | Parameters | Notes |
|--------|-----------|-------|
| **SOLO** | `-APIKEY`, `-IOC_FilePath` | Scans all indicator types separately |
| **MIXR** | `-APIKEY`, `-IOC_FilePath` | Requires Excel installed |
| **AbuseIPDB** | `-APIKEY`, `-FilePath` | Only processes IP addresses |

## 📊 Output Examples

### SOLO Output
- Individual CSV or Excel files for each indicator type:
  - `Hash_Report.xlsx`
  - `Domain_Report.xlsx`
  - `IP_Report.xlsx`
  - `Url_Report.xlsx`

### MIXR Output
- Single consolidated Excel workbook (e.g., `Threat_Analysis_Report.xlsx`) with:
  - Sheet 1: Hash Analysis
  - Sheet 2: Domain Analysis
  - Sheet 3: IP Analysis
  - Sheet 4: URL Analysis

### AbuseIPDB Output
- Excel or CSV report (e.g., `IP_Report.xlsx`) containing:
  - **IPAddress**: The queried IP address
  - **ISP**: Internet Service Provider information
  - **TotalReports**: Number of abuse reports filed
  - **AbuseConfidenceScore**: Trust score (0-100%)
  - **Domain**: Associated domain name
  - **Whitelisted**: Whether the IP is whitelisted
  - **IsTor**: Tor exit node status
  - **UsageType**: Classification (e.g., "hosting", "residential")
  - **CountryCode**: 2-letter country code
  - **CountryName**: Full country name

## ⚠️ Important Notes

### Error Handling
- Invalid indicators are logged with warnings (e.g., malformed hashes, invalid IPs)
- Indicators not found in VirusTotal are reported with 404 status
- API rate limits are handled gracefully with retry logic
- AbuseIPDB rate limit hits (429 status) stop further requests with a warning

### Excel Requirement
- **MIXR** requires Microsoft Excel to be installed (validated via registry check)
- **SOLO** does NOT require Excel
- **AbuseIPDB** does NOT require Excel
- If Excel is not installed, MIXR will display an error message

### API Rate Limiting
- **VirusTotal API**: Rate limits depend on your subscription tier
- **AbuseIPDB API**: Has its own rate limiting (check your subscription tier)
- Both suites handle throttling automatically
- Consider batching large IOC files to avoid hitting rate limits
- AbuseIPDB will stop processing and log a warning when rate limit is hit (HTTP 429)

### File Encoding
- IOC files should be encoded in UTF-8 or ANSI format
- One indicator per line
- Blank lines are skipped automatically
- Mixed indicator types are supported (AbuseIPDB automatically filters for IPs only)

### IP Validation
- AbuseIPDB validates both IPv4 and IPv6 addresses using strict regex patterns
- Invalid IP formats are silently skipped
- De-duplication ensures each unique IP is queried only once

## 🔧 Advanced Usage

### Batch Processing Multiple Files

```powershell
# VirusTotal MIXR - Multiple files
$ApiKey = "your_virustotal_api_key"
$IocFiles = Get-ChildItem -Path "C:\IOCs" -Filter "*.txt"

foreach ($File in $IocFiles) {
    .\MIXR.ps1 -APIKEY $ApiKey -IOC_FilePath $File.FullName
}
```

```powershell
# AbuseIPDB - Multiple IP files
$AbuseIPDBKey = "your_abuseipdb_api_key"
$IpFiles = Get-ChildItem -Path "C:\IPs" -Filter "*.txt"

foreach ($File in $IpFiles) {
    .\AbuseIPDB.ps1 -APIKEY $AbuseIPDBKey -FilePath $File.FullName
}
```

### Combined Analysis Workflow

```powershell
# 1. First, run AbuseIPDB to get detailed IP intelligence
$AbuseIPDBKey = "your_abuseipdb_api_key"
.\AbuseIPDB.ps1 -APIKEY $AbuseIPDBKey -FilePath "C:\Threats\all_iocs.txt"

# 2. Then, run MIXR for comprehensive multi-vector analysis
$VTKey = "your_virustotal_api_key"
.\MIXR.ps1 -APIKEY $VTKey -IOC_FilePath "C:\Threats\all_iocs.txt"
```

### Comparing SOLO vs MIXR vs AbuseIPDB

```powershell
# For rapid feedback on individual indicator types
.\SOLO.ps1 -APIKEY $VTKey -IOC_FilePath "C:\Threats\iocs.txt"

# For consolidated reporting
.\MIXR.ps1 -APIKEY $VTKey -IOC_FilePath "C:\Threats\iocs.txt"

# For deep-dive IP reputation analysis
.\AbuseIPDB.ps1 -APIKEY $AbuseIPDBKey -FilePath "C:\Threats/iocs.txt"
```

### Storing API Keys Securely (PowerShell Profile)

Add to your PowerShell profile (`$PROFILE`):
```powershell
$env:VT_API_KEY = "your_virustotal_api_key"
$env:ABUSEIPDB_API_KEY = "your_abuseipdb_api_key"
```

Or use secure prompting:
```powershell
$env:VT_API_KEY = Read-Host -AsSecureString -Prompt "Enter VirusTotal API Key"
$env:ABUSEIPDB_API_KEY = Read-Host -AsSecureString -Prompt "Enter AbuseIPDB API Key"
```

Then use in scripts:
```powershell
.\MIXR.ps1 -APIKEY $env:VT_API_KEY -IOC_FilePath "C:\Threats\iocs.txt"
.\AbuseIPDB.ps1 -APIKEY $env:ABUSEIPDB_API_KEY -FilePath "C:\Threats\ips.txt"
```

## 📝 Supported File Formats for IOCs

- **Plain Text (.txt)**: One indicator per line
- **Comma-Separated (.csv)**: One indicator per line (commas ignored)
- **Relative or Absolute Paths**: Both are supported

**Example IOC file (iocs.txt):**
```
4d967a2a964760d93477d2d023efb0b1
a94a8fe5ccb19ba61c4c0873d391e987982fbbd3
example.com
192.0.2.1
https://example.com/payload.exe
malicious.net
::1
badactor.io
```

## 🐛 Troubleshooting

### "Microsoft Excel is not installed on this system"
- **Issue**: Trying to run MIXR without Excel
- **Solution**: Either install Excel or use SOLO instead for separate reports

### "Invalid API Key"
- **Issue**: VirusTotal or AbuseIPDB API returns 401 Unauthorized
- **Solution**: Verify your API key is correct and active at the respective service website

### "File not found" Error
- **Issue**: `-IOC_FilePath` or `-FilePath` points to a non-existent file
- **Solution**: Use absolute paths or ensure the relative path is correct from the script directory

### Rate Limit Errors (429)
- **Issue**: VirusTotal or AbuseIPDB API throttling due to too many requests
- **Solution**: Wait a moment and retry, or upgrade your subscription tier
- **AbuseIPDB Specific**: The script will automatically stop processing and log a warning

### "No data captured to create a report"
- **Issue**: No valid indicators of the expected type were found
- **Causes**:
  - IOC file is empty
  - All indicators are malformed or don't match the expected format
  - For AbuseIPDB: No valid IPv4/IPv6 addresses found in the file
- **Solution**: Verify your IOC file contains properly formatted indicators

### Mixed Indicator Type Processing
- **SOLO/MIXR**: Will process all indicator types and skip malformed entries
- **AbuseIPDB**: Only processes IPv4 and IPv6 addresses; other indicator types are silently filtered out
- **Recommendation**: Use AbuseIPDB with mixed IOC files if you only need IP intelligence

## 🤝 Contributing

Contributions are welcome! If you'd like to add support for additional indicator types or improve the reporting engine:

1. Fork the repository
2. Create a feature branch
3. To add a new scanner module:
   - Create a new PowerShell script in the `virustotal/` directory
   - Implement your scanner class with the same structure as existing scanners (e.g., `Hash.ps1`)
   - Update SOLO.ps1 and MIXR.ps1 to include your new module in the `$Scanners` manifest
   - Test with sample IOCs
4. To add a new standalone analysis tool:
   - Create a new PowerShell script in the root directory
   - Follow the same parameter structure as SOLO.ps1/MIXR.ps1
   - Utilize shared utilities from `root/` directory
5. Submit a pull request

### Development Guidelines
- Use type-safe .NET collections (HashSet, List) for data structures
- Implement proper error handling for API responses
- Validate input indicators using regex patterns
- Add helpful comments to complex logic
- Test with both valid and invalid IOCs

## 📄 License

This project is maintained by **SivaMani70** (May 2026)

## 🙏 Acknowledgments

- Built for the security community to streamline malware analysis
- Powered by [VirusTotal's API v3](https://developers.virustotal.com/reference/)
- Designed for enterprise threat intelligence workflows

## 📞 Support

For issues, feature requests, or questions:
- Check the troubleshooting section above
- Review your API key configuration
- Ensure all IOCs are in a valid format
- Verify your network can reach VirusTotal's API

---

**Last Updated**: May 2026  
**Version**: 1.0  
**Maintained by**: SivaMani70

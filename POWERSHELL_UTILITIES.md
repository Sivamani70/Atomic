# PowerShell Utility Scripts

This document provides an overview of the helper PowerShell scripts that complement the core threat intelligence analysis tools (SOLO, MIXR, AbuseIPDB). These utilities perform essential data processing, transformation, and utility functions for the threat intelligence workflow.

## 📑 Table of Contents

- [PowerShell Utility Scripts](#powershell-utility-scripts)
  - [📑 Table of Contents](#-table-of-contents)
  - [TMC.ps1](#tmcps1)
    - [Purpose](#purpose)
    - [Key Features](#key-features)
    - [Syntax](#syntax)
    - [Parameters](#parameters)
    - [Examples](#examples)
    - [Requirements](#requirements)
    - [Notes](#notes)
  - [IOC_90Day_Aggregator.ps1](#ioc_90day_aggregatorps1)
    - [Purpose](#purpose-1)
    - [Key Features](#key-features-1)
    - [Syntax](#syntax-1)
    - [Parameters](#parameters-1)
    - [Examples](#examples-1)
    - [Features](#features)
    - [Requirements](#requirements-1)
    - [Notes](#notes-1)
  - [separator.ps1](#separatorps1)
    - [Purpose](#purpose-2)
    - [Key Features](#key-features-2)
    - [Syntax](#syntax-2)
    - [Parameters](#parameters-2)
    - [Examples](#examples-2)
    - [Output](#output)
    - [IOC Types Recognized](#ioc-types-recognized)
    - [Requirements](#requirements-2)
    - [Notes](#notes-2)
  - [TimeConverter.ps1](#timeconverterps1)
    - [Purpose](#purpose-3)
    - [Key Features](#key-features-3)
    - [Syntax](#syntax-3)
    - [Supported Time Zones](#supported-time-zones)
    - [Example Usage](#example-usage)
    - [Input Format](#input-format)
    - [Daylight Saving Time Handling](#daylight-saving-time-handling)
    - [Requirements](#requirements-3)
    - [Notes](#notes-3)
  - [Integration with Core Tools](#integration-with-core-tools)
  - [Common Workflows](#common-workflows)
    - [Workflow 1: Process Raw IOC List to Excel Report](#workflow-1-process-raw-ioc-list-to-excel-report)
    - [Workflow 2: 90-Day Threat Intelligence Review](#workflow-2-90-day-threat-intelligence-review)
    - [Workflow 3: International Log Timeline Analysis](#workflow-3-international-log-timeline-analysis)
  - [Author](#author)
  - [See Also](#see-also)

## TMC.ps1

**Enterprise-Grade Threat Intelligence Consolidation Engine**

### Purpose

The TMC (Threat Management Consolidation) script is an advanced data pipeline tool designed to consolidate, cleanse, and map multi-source threat intelligence data into formatted Excel reports. It acts as a specialized Threat Intelligence parsing engine that combines raw IOC data from multiple workbooks into deduplicated, normalized, and structured Excel reports.

### Key Features

- **Multi-Source Ingestion**: Accepts multiple Excel workbooks as input sources
- **Data Cleansing**: Removes noise and sanitizes malware actor naming structures
- **Deduplication**: Uses high-performance .NET `HashSet` structures for efficient deduplication
- **Excel COM Automation**: Low-level COM automation with Microsoft Excel for native format support
- **STIX-Like Reporting**: Generates schema-compliant structured reports
- **Metrics Telemetry**: Provides detailed counts of all IOC types processed

### Syntax

```powershell
.\TMC.ps1 -WorkBooks <string[]>
```

### Parameters

| Parameter   | Type     | Required | Description                                            |
| ----------- | -------- | -------- | ------------------------------------------------------ |
| `WorkBooks` | string[] | Yes      | Array of absolute paths to source Excel workbook files |

### Examples

**Single workbook:**

```powershell
.\TMC.ps1 -WorkBooks "D:\IOCs\ActorGroupA.xlsx"
```

**Multiple workbooks:**

- Do not add space b/w comma (,)

```powershell
.\TMC.ps1 -WorkBooks "D:\IOCs\ActorGroupA.xlsx","D:\IOCs\CampaignB.xlsx","D:\IOCs\CampaignC.xlsx"
```

### Requirements

- Microsoft Excel installed locally (COM Interop validation executed against HKLM registry)
- Access to the specified workbook files
- Write permissions for output report generation

### Notes

- Implements programmatic Marshal COM reference release wrappers to prevent zombie `excel.exe` processes
- Supports all indicator types: MD5, SHA1, SHA256, Domains, URLs, Emails, and IP addresses
- Output files are typically saved in a structured format with deduplication metadata

---

## IOC_90Day_Aggregator.ps1

**90-Day Indicator of Compromise Aggregation Engine**

### Purpose

The IOC_90Day_Aggregator script specializes in collecting, processing, and aggregating Indicators of Compromise (IOCs) from multiple sources over a rolling 90-day period. This script is designed for organizations that need to maintain a comprehensive, time-windowed view of threat intelligence data.

### Key Features

- **Time-Windowed Aggregation**: Focuses on IOCs from the last 90 days
- **Multi-Workbook Processing**: Ingests data from multiple Excel files simultaneously
- **Data Normalization**: Standardizes all indicators for consistency
- **Deduplication**: Eliminates duplicate entries using .NET `HashSet` structures
- **Performance Optimization**: O(1) lookup complexity for large-volume indicator handling
- **Enterprise-Grade Processing**: Implements high-performance data structures for scalability

### Syntax

```powershell
.\IOC_90Day_Aggregator.ps1 -InputDirectory <string> -MonthName <string>
```

### Parameters

| Parameter        | Type   | Required | Description                                          |
| ---------------- | ------ | -------- | ---------------------------------------------------- |
| `InputDirectory` | string | Yes      | Directory containing source workbook files to ingest |
| `MonthName`      | string | Yes      | Name of the month for which to generate reports      |

### Examples

**January aggregation:**

```powershell
.\IOC_90Day_Aggregator.ps1 -InputDirectory "D:\IOCs\Quarterly\Q1" -MonthName "January"
```

**Current month aggregation:**

```powershell
.\IOC_90Day_Aggregator.ps1 -InputDirectory "D:\ThreatIntel\Current" -MonthName "July"
```

### Features

- Processes multiple IOC types: MD5, SHA1, SHA256, Domains, URLs, Emails, IP addresses
- Generates metrics and telemetry for reporting purposes
- Maintains compatibility with Excel COM automation for native format output
- Eliminates zombie Excel processes through programmatic COM cleanup

### Requirements

- Access to the input directory with workbook files
- Microsoft Excel installed locally
- Sufficient system resources for processing large IOC volumes

### Notes

- Particularly useful for compliance reporting and historical threat analysis
- Supports rolling 90-day windows for ongoing threat monitoring
- Aggregates data from multiple campaigns or threat actors

---

## separator.ps1

**IOC Parser, Normalizer, and Categorizer**

### Purpose

The separator script is a comprehensive IOC processing utility that reads raw, unformatted lists of Indicators of Compromise and transforms them into clean, categorized, deduplicated output. It automatically reverses common sanitization techniques and categorizes indicators by type.

### Key Features

- **Sanitization Reversal**: Automatically restores commonly sanitized indicators
  - Converts `[.]` back to `.`
  - Converts `[:]` back to `:`
  - Handles other common obfuscation patterns
- **Normalization**: Converts all entries to lowercase and trims whitespace
- **Automatic Categorization**: Uses Regular Expressions to categorize IOCs
- **Deduplication**: Leverages .NET `HashSet` for efficient duplicate removal
- **Manual Review Flagging**: Marks unrecognized or unparseable items for manual review

### Syntax

```powershell
.\separator.ps1 -Source <string>
```

### Parameters

| Parameter | Type   | Required | Description                                   |
| --------- | ------ | -------- | --------------------------------------------- |
| `Source`  | string | Yes      | Path to the text file containing raw IOC list |

### Examples

**Process raw IOC file:**

```powershell
.\separator.ps1 -Source "C:\ThreatIntel\raw_ioc.txt"
```

**Process from alternate location:**

```powershell
.\separator.ps1 -Source "D:\Downloads\suspected_malware_list.txt"
```

### Output

Generates a file named `IOCs.txt` in the current working directory containing:

- **Categorized IOCs**: Organized by type (MD5, SHA1, SHA256, Domains, URLs, Emails, IPs)
- **Deduplication**: No duplicate entries
- **Flagged Items**: Unrecognized entries marked for manual review

### IOC Types Recognized

| Type           | Format                    | Example                                                            |
| -------------- | ------------------------- | ------------------------------------------------------------------ |
| **MD5**        | 32 hexadecimal characters | `4d967a2a964760d93477d2d023efb0b1`                                 |
| **SHA1**       | 40 hexadecimal characters | `a94a8fe5ccb19ba61c4c0873d391e987982fbbd3`                         |
| **SHA256**     | 64 hexadecimal characters | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| **Domain**     | FQDN format               | `malicious.com`                                                    |
| **URL**        | HTTP/HTTPS URL            | `https://example.com/payload.exe`                                  |
| **Email**      | Email address format      | `attacker@malicious.com`                                           |
| **IP Address** | IPv4 or IPv6              | `192.0.2.1` or `2001:db8::1`                                       |

### Requirements

- Read access to the source file
- Write permissions in the current working directory
- No external dependencies

### Notes

- Perfect for processing threat feeds and raw IOC lists
- Handles mixed indicator types in a single input file
- Output file format is human-readable text for easy review
- Use the output file directly with SOLO, MIXR, or AbuseIPDB scripts

---

## TimeConverter.ps1

**Interactive Time Zone Conversion Utility**

### Purpose

The TimeConverter script provides an interactive utility to convert dates and times between multiple supported time zones. It leverages the Windows .NET `System.TimeZoneInfo` class for accurate conversions, including automatic Daylight Saving Time (DST) adjustments.

### Key Features

- **Multi-Zone Support**: Converts between IST, UTC, MST, and EST time zones
- **Automatic DST Adjustment**: Handles Daylight Saving Time conversion automatically (EDT for EST, MDT for MST)
- **Interactive Menu**: User-friendly command-line interface for time zone selection
- **Native Windows Support**: Uses Windows Time Zone IDs for accuracy
- **Custom PowerShell Class**: Implements a specialized `Convert` class for time zone operations

### Syntax

```powershell
.\TimeConverter.ps1
```

### Supported Time Zones

| Option | Time Zone                        | Notes                              |
| ------ | -------------------------------- | ---------------------------------- |
| 1      | India Standard Time (IST)        | UTC+5:30                           |
| 2      | Coordinated Universal Time (UTC) | UTC±0                              |
| 3      | Mountain Standard Time (MST)     | UTC-7 (MST) / UTC-6 (MDT with DST) |
| 4      | Eastern Standard Time (EST)      | UTC-5 (EST) / UTC-4 (EDT with DST) |

### Example Usage

```powershell
.\TimeConverter.ps1

Available Time Zones to Convert
1. India Standard Time
2. UTC
3. Mountain Standard Time
4. Eastern Standard Time

Choose the Input TimeZone [b/w 1 - 4]: 4
Choose the Output TimeZone [b/w 1 - 4]: 1
Enter Date Time to convert [DD/MM/YYYY HH:MM:SS]: 07/25/2026 14:00:00

Converted Time: 07/26/2026 00:30:00
```

### Input Format

- **Date/Time Format**: `DD/MM/YYYY HH:MM:SS`
- **Example**: `07/25/2026 14:00:00` represents July 25, 2026 at 2:00 PM

### Daylight Saving Time Handling

The script automatically adjusts for DST:

- **EST to EDT**: Automatically converts to EDT (UTC-4) during daylight saving period (typically March-November)
- **MST to MDT**: Automatically converts to MDT (UTC-6) during daylight saving period

Example with DST:

- **Input**: July 15, 2026 14:00:00 EST
- **Automatic Adjustment**: Script recognizes July is in DST period and treats it as EDT (UTC-4)
- **Output**: Converts accurately based on the current DST rules for 2026

### Requirements

- Windows Time Zone IDs properly configured
- No external dependencies
- Fully interactive (no batch processing)

### Notes

- Designed for interactive use; does not accept command-line parameters
- Ideal for quick time zone conversions during incident response or international analysis
- Windows automatically handles Daylight Saving Time transitions
- May be useful when correlating logs or evidence from different geographic regions

---

## Integration with Core Tools

These utility scripts work seamlessly with the main threat intelligence analysis tools:

| Utility                      | Works With             | Purpose                                        |
| ---------------------------- | ---------------------- | ---------------------------------------------- |
| **separator.ps1**            | SOLO, MIXR, AbuseIPDB  | Pre-process raw IOC lists before analysis      |
| **TMC.ps1**                  | MIXR (Excel output)    | Consolidate MIXR results across multiple runs  |
| **IOC_90Day_Aggregator.ps1** | MIXR (historical data) | Aggregate historical IOCs for trend analysis   |
| **TimeConverter.ps1**        | All tools              | Convert timestamps for timeline reconstruction |

---

## Common Workflows

### Workflow 1: Process Raw IOC List to Excel Report

1. Use **separator.ps1** to clean and categorize raw IOC list
2. Use **SOLO.ps1** or **MIXR.ps1** to analyze categorized IOCs
3. Use **TMC.ps1** to consolidate multiple analysis reports

### Workflow 2: 90-Day Threat Intelligence Review

1. Collect IOC files from multiple months in a directory
2. Run **IOC_90Day_Aggregator.ps1** to aggregate and deduplicate
3. Use **MIXR.ps1** on aggregated data for comprehensive analysis
4. Review results for trends and threat actor patterns

### Workflow 3: International Log Timeline Analysis

1. Collect logs/evidence from multiple geographic regions
2. Use **TimeConverter.ps1** to normalize all timestamps to UTC
3. Create timeline for correlation analysis

---

## Author

These utility scripts were created by **SivaMani70** and are maintained as part of the enterprise threat intelligence toolkit (May-July 2026).

## See Also

- [Main README](README.md) - Core tool documentation
- [SOLO.ps1](SOLO.ps1) - Single-engine IOC analyzer
- [MIXR.ps1](MIXR.ps1) - Multi-engine IOC analyzer with Excel reporting
- [AbuseIPDB.ps1](AbuseIPDB.ps1) - IP reputation analyzer

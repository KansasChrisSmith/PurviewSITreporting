# Microsoft Purview SharePoint Sensitive Information Type Reporting

## Overview

This PowerShell script provides a streamlined way to export Microsoft Purview Sensitive Information Type (SIT) findings for SharePoint Online.

Rather than requiring administrators to maintain separate input files for SharePoint sites and Sensitive Information Types, the script dynamically retrieves the relevant SIT definitions and queries Purview Content Explorer for matching findings.

The script only includes SITs with actual findings in the reporting output. It also supports pagination, session reconnection, retry handling, error logging, and separate detailed and summary CSV exports.

> Important: This script does not scan SharePoint files directly. It queries classification findings already available through Microsoft Purview Content Explorer.

## What the Script Does

At a high level, the script performs the following workflow:

1. Connects to Microsoft Purview by using `Connect-IPPSSession`.
2. Retrieves the available Sensitive Information Types by using `Get-DlpSensitiveInformationType`.
3. Iterates through the discovered SITs.
4. Queries Content Explorer by using `Export-ContentExplorerData`.
5. Restricts the search to the SharePoint Online (SPO) workload.
6. Processes all result pages by using the returned page cookie.
7. Extracts the relevant SIT confidence details from each finding.
8. Writes actual SharePoint findings to a detailed CSV file.
9. Creates a separate SIT-level summary CSV file.
10. Logs failures or incomplete exports to a dedicated error CSV file.
11. Reconnects the Purview session periodically during long-running exports.
12. Retries failed Content Explorer requests before marking a SIT as incomplete.

## Reporting Behavior

The script does not create reporting rows simply because a SIT exists.

A detailed reporting row is written only when `Export-ContentExplorerData` returns a SharePoint finding for that SIT. This keeps the detailed CSV focused on actual detected content.

For example, if Content Explorer contains detections for:

- Indonesia Passport Number in Exchange
- U.S. Social Security Number in SharePoint
- Credit Card Number in SharePoint

The script's `Workload = "SPO"` scope means the Exchange-only Indonesia Passport Number findings are excluded, while qualifying SharePoint findings are retained.

This behavior is intentional and keeps the output aligned to the SharePoint reporting scope.

## Output Files

The script generates the following CSV files in the configured output folder:

| File | Description |
| --- | --- |
| `DetailedReport.csv` | Contains individual SharePoint findings for SITs with actual detections |
| `SummaryReport.csv` | Provides a SIT-level summary of findings, confidence totals, and processing status |
| `ErrorReport.csv` | Captures failed or incomplete exports and associated error details |

## Summary Report Fields

The SIT summary output contains the following fields:

| Field | Description |
| --- | --- |
| `SITName` | Sensitive Information Type |
| `SITId` | SIT identifier |
| `FindingsReturned` | Number of findings processed for the SIT |
| `LowConfidenceTotal` | Total low-confidence matches |
| `MediumConfidenceTotal` | Total medium-confidence matches |
| `HighConfidenceTotal` | Total high-confidence matches |
| `PagesProcessed` | Number of Content Explorer pages processed |
| `Status` | `Complete` or `Incomplete` |

## Parameters

The script supports the following parameters:

| Parameter | Default | Purpose |
| --- | ---: | --- |
| `OutFolder` | `./PurviewResults` | Location for generated CSV reports |
| `PageSize` | `10000` | Page size requested from Content Explorer |
| `MaxRetries` | `5` | Maximum attempts for a failed export request |
| `ReconnectMinutes` | `50` | Interval at which the Purview session is recreated |

## Notes

- This solution is intended for Microsoft Purview reporting scenarios focused on SharePoint Online content.
- Output is limited to actual findings returned by Content Explorer for the selected workload.
- The script includes built-in retry logic and session refresh behavior to improve long-running export reliability.

## Related Use Case

This script is useful for administrators who need to quickly identify which SharePoint Online SITs are producing actual detections and summarize those results in a simple CSV-based reporting format.

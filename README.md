Microsoft Purview SharePoint Sensitive Information Type Reporting
Overview

This PowerShell script provides a streamlined way to export Microsoft Purview Sensitive Information Type (SIT) findings for SharePoint Online.

Rather than requiring administrators to maintain separate input files containing SharePoint sites and Sensitive Information Types, the script dynamically retrieves the Sensitive Information Types available through Purview and checks each SIT for findings in the SharePoint Online (SPO) workload.

Only SITs with actual findings are included in the reporting output. The script also includes pagination, session reconnection, retry handling, error logging, and separate detailed and summary CSV reports.

Important: This script does not scan SharePoint files for sensitive information itself. It queries classification findings available through Microsoft Purview Content Explorer. Classification and discovery of the underlying content must already have occurred through Purview.

What the Script Does

At a high level, the script performs the following workflow:

Connects to Microsoft Purview using Connect-IPPSSession.
Dynamically retrieves Sensitive Information Types using Get-DlpSensitiveInformationType.
Iterates through the discovered SITs.
Queries Content Explorer using Export-ContentExplorerData.
Restricts the search to the SharePoint Online (SPO) workload.
Processes all available result pages using the returned page cookie.
Extracts the matching SIT confidence information from each finding.
Writes actual SharePoint findings to a detailed CSV.
Creates a separate SIT-level summary CSV.
Logs failures or incomplete exports to a dedicated error CSV.
Periodically reconnects the Purview PowerShell session during long-running exports.
Retries failed Content Explorer requests before marking a SIT incomplete.

Reports Actual Purview Findings

The script does not create reporting rows simply because a SIT exists.

A detailed reporting row is written only when Export-ContentExplorerData returns a SharePoint finding for that SIT.

This keeps the detailed CSV focused on actual detected content.

For example, if Content Explorer contains detections for:

Indonesia Passport Number in Exchange
U.S. Social Security Number in SharePoint
Credit Card Number in SharePoint

the script's Workload = "SPO" scope means the Exchange-only Indonesia Passport Number findings are not included, while qualifying SharePoint findings can be returned.

This distinction is intentional.


| Field                   | Description                                |
| ----------------------- | ------------------------------------------ |
| `SITName`               | Sensitive Information Type                 |
| `SITId`                 | SIT identifier                             |
| `FindingsReturned`      | Number of findings processed for the SIT   |
| `LowConfidenceTotal`    | Total low-confidence matches               |
| `MediumConfidenceTotal` | Total medium-confidence matches            |
| `HighConfidenceTotal`   | Total high-confidence matches              |
| `PagesProcessed`        | Number of Content Explorer pages processed |
| `Status`                | `Complete` or `Incomplete`                 |


| Parameter          |            Default | Purpose                                            |
| ------------------ | -----------------: | -------------------------------------------------- |
| `OutFolder`        | `.\PurviewResults` | Location for generated CSV reports                 |
| `PageSize`         |            `10000` | Page size requested from Content Explorer          |
| `MaxRetries`       |                `5` | Maximum attempts for a failed export request       |
| `ReconnectMinutes` |               `50` | Interval at which the Purview session is recreated |

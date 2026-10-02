<# 
"metadata": {
    "description": "SIT Reporting – Script reviews all active SITs in Purvew for positive hits and exports output to csv files",
    "disclaimer": "THIS SAMPLE CODE IS PROVIDED FOR THE PURPOSE OF ILLUSTRATION ONLY AND IS NOT INTENDED TO BE USED IN A PRODUCTION ENVIRONMENT. THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED \"AS IS\" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.\n\nMicrosoft grants you a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that you agree:\n(i) not to use Microsoft’s name, logo, or trademarks to market your software product in which the Sample Code is embedded;\n(ii) to include a valid copyright notice on your software product in which the Sample Code is embedded; and\n(iii) to indemnify, hold harmless, and defend Microsoft and its suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code."
  }
#>





[CmdletBinding()]
param(
    [string]$OutFolder = ".\PurviewResults",
    [ValidateRange(1, 10000)]
    [int]$PageSize = 10000,
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 5,
    [ValidateRange(5, 55)]
    [int]$ReconnectMinutes = 50
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $OutFolder)) {
    New-Item -Path $OutFolder -ItemType Directory -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$detailCsv = Join-Path $OutFolder "Purview_SPO_SIT_Findings_$timestamp.csv"
$summaryCsv = Join-Path $OutFolder "Purview_SPO_SIT_Summary_$timestamp.csv"
$errorCsv = Join-Path $OutFolder "Purview_SPO_SIT_Errors_$timestamp.csv"

$detailCsvInitialized = $false
$errorCsvInitialized = $false

Import-Module ExchangeOnlineManagement -ErrorAction Stop

function Connect-PurviewSession {
    Write-Host "Connecting to Microsoft Purview PowerShell..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Connect-IPPSSession -ErrorAction Stop
    return Get-Date
}

function Get-MatchedSITConfidence {
    param(
        [Parameter(Mandatory)]
        [object]$Record,

        [Parameter(Mandatory)]
        [string]$SitId
    )

    $result = [ordered]@{
        LowConfidenceMatch    = 0
        MediumConfidenceMatch = 0
        HighConfidenceMatch   = 0
    }

    $sitJson = [string]$Record.SensitiveInfoTypesData
    if ([string]::IsNullOrWhiteSpace($sitJson)) {
        return [pscustomobject]$result
    }

    try {
        $sitData = $sitJson | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        return [pscustomobject]$result
    }

    $matchedSit = @($sitData) |
        Where-Object { [string]$_.Id -eq $SitId } |
        Select-Object -First 1

    if ($null -eq $matchedSit) {
        return [pscustomobject]$result
    }

    foreach ($property in @(
        "LowConfidenceMatch",
        "MediumConfidenceMatch",
        "HighConfidenceMatch"
    )) {
        if ($null -ne $matchedSit.$property) {
            $result[$property] = [int]$matchedSit.$property
        }
    }

    return [pscustomobject]$result
}

function Write-ExportError {
    param(
        [string]$SitName,
        [string]$SitId,
        [int]$PageNumber,
        [string]$PageCookie,
        [string]$Message
    )

    $errorRow = [pscustomobject]@{
        Timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        SITName    = $SitName
        SITId      = $SitId
        PageNumber = $PageNumber
        PageCookie = $PageCookie
        Error      = $Message
    }

    if ($script:errorCsvInitialized) {
        $errorRow | Export-Csv -Path $errorCsv -NoTypeInformation -Encoding UTF8 -Append
    }
    else {
        $errorRow | Export-Csv -Path $errorCsv -NoTypeInformation -Encoding UTF8
        $script:errorCsvInitialized = $true
    }
}

$connectedAt = Connect-PurviewSession

Write-Host "Discovering Sensitive Information Types..." -ForegroundColor Cyan
$sits = @(
    Get-DlpSensitiveInformationType |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Name) } |
        Sort-Object Name
)

if ($sits.Count -eq 0) {
    throw "No Sensitive Information Types were returned by Purview."
}

Write-Host "Discovered $($sits.Count) Sensitive Information Types." -ForegroundColor Green

$summary = [System.Collections.Generic.List[object]]::new()
$currentSitNumber = 0

foreach ($sit in $sits) {
    $currentSitNumber++
    $sitName = [string]$sit.Name
    $sitId = [string]$sit.Id

    if ([string]::IsNullOrWhiteSpace($sitName)) {
        continue
    }

    Write-Host ""
    Write-Host "[$currentSitNumber/$($sits.Count)] Processing SIT: $sitName" -ForegroundColor Cyan

    $pageCookie = $null
    $pageNumber = 0
    $sitFindingCount = 0
    $sitLowCount = 0
    $sitMediumCount = 0
    $sitHighCount = 0
    $sitFailed = $false

    do {
        $pageNumber++

        if (((Get-Date) - $connectedAt).TotalMinutes -ge $ReconnectMinutes) {
            $connectedAt = Connect-PurviewSession
        }

        $response = $null
        $requestSucceeded = $false
        $lastError = $null

        for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
            try {
                Write-Host "Requesting page $pageNumber, attempt $attempt..." -ForegroundColor DarkGray

                $parameters = @{
                    TagType      = "SensitiveInformationType"
                    TagName      = $sitName
                    Workload     = "SPO"
                    PageSize     = $PageSize
                    WarningAction = "SilentlyContinue"
                    ErrorAction  = "Stop"
                }

                if (-not [string]::IsNullOrWhiteSpace([string]$pageCookie)) {
                    $parameters.PageCookie = $pageCookie
                }

                $response = @(Export-ContentExplorerData @parameters)
                $requestSucceeded = $true
                break
            }
            catch {
                $lastError = $_.Exception.Message
                Write-Warning "SIT '$sitName', page $pageNumber, attempt $attempt failed: $lastError"

                if ($attempt -lt $MaxRetries) {
                    try {
                        $connectedAt = Connect-PurviewSession
                    }
                    catch {
                        Write-Warning "Purview reconnection failed: $($_.Exception.Message)"
                    }

                    $delaySeconds = [int][math]::Min(300, ([math]::Pow(2, $attempt) * 10))
                    Write-Host "Waiting $delaySeconds seconds before retry..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $delaySeconds
                }
            }
        }

        if (-not $requestSucceeded) {
            Write-ExportError -SitName $sitName -SitId $sitId -PageNumber $pageNumber -PageCookie $pageCookie -Message $lastError
            Write-Warning "Skipping the remainder of SIT '$sitName' after $MaxRetries failed attempts."
            $sitFailed = $true
            break
        }

        if ($response.Count -eq 0) {
            Write-Host "No findings returned." -ForegroundColor DarkGray
            break
        }

        $metadata = $response[0]
        $records = @()

        if ($response.Count -gt 1) {
            $records = @($response | Select-Object -Skip 1)
        }

        if ($records.Count -gt 0) {
            $outputRows = foreach ($record in $records) {
                $confidence = Get-MatchedSITConfidence -Record $record -SitId $sitId

                $sitFindingCount++
                $sitLowCount += $confidence.LowConfidenceMatch
                $sitMediumCount += $confidence.MediumConfidenceMatch
                $sitHighCount += $confidence.HighConfidenceMatch

                [pscustomobject]@{
                    SITName               = $sitName
                    SITId                 = $sitId
                    FileName              = $record.FileName
                    FileSourceUrl         = $record.FileSourceUrl
                    LowConfidenceMatch    = $confidence.LowConfidenceMatch
                    MediumConfidenceMatch = $confidence.MediumConfidenceMatch
                    HighConfidenceMatch   = $confidence.HighConfidenceMatch
                    SensitiveInfoTypesData = $record.SensitiveInfoTypesData
                }
            }

            if ($detailCsvInitialized) {
                $outputRows | Export-Csv -Path $detailCsv -NoTypeInformation -Encoding UTF8 -Append
            }
            else {
                $outputRows | Export-Csv -Path $detailCsv -NoTypeInformation -Encoding UTF8
                $detailCsvInitialized = $true
            }
        }

        $recordsReturned = $records.Count
        if ($null -ne $metadata.RecordsReturned) {
            $recordsReturned = [int]$metadata.RecordsReturned
        }

        $morePages = ([string]$metadata.MorePagesAvailable -eq "True")
        $pageCookie = [string]$metadata.PageCookie

        Write-Host "Page ${pageNumber}: $recordsReturned finding(s); more pages: $morePages" -ForegroundColor Gray

        if ($morePages -and [string]::IsNullOrWhiteSpace($pageCookie)) {
            Write-ExportError -SitName $sitName -SitId $sitId -PageNumber $pageNumber -PageCookie $pageCookie -Message "Purview reported additional pages but did not return a page cookie."
            $sitFailed = $true
            break
        }
    }
    while ($morePages)

    if ($sitFindingCount -gt 0 -or $sitFailed) {
        $status = if ($sitFailed) { "Incomplete" } else { "Complete" }

        $summary.Add([pscustomobject]@{
            SITName               = $sitName
            SITId                 = $sitId
            FindingsReturned      = $sitFindingCount
            LowConfidenceTotal    = $sitLowCount
            MediumConfidenceTotal = $sitMediumCount
            HighConfidenceTotal   = $sitHighCount
            PagesProcessed        = $pageNumber
            Status                = $status
        })
    }

    if ($sitFindingCount -gt 0) {
        Write-Host "Completed '$sitName': $sitFindingCount finding(s)." -ForegroundColor Green
    }
    elseif (-not $sitFailed) {
        Write-Host "Completed '$sitName': no SharePoint findings." -ForegroundColor DarkGray
    }
}

if ($summary.Count -gt 0) {
    $summary |
        Sort-Object SITName |
        Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8
}

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Purview export completed." -ForegroundColor Green

if ($detailCsvInitialized) {
    Write-Host "Detailed findings: $detailCsv" -ForegroundColor Green
}
else {
    Write-Host "Purview returned no SharePoint findings. No detailed findings CSV was created." -ForegroundColor Yellow
}

if ($summary.Count -gt 0) {
    Write-Host "Summary: $summaryCsv" -ForegroundColor Green
}

if ($errorCsvInitialized) {
    Write-Host "Errors or incomplete exports: $errorCsv" -ForegroundColor Yellow
}
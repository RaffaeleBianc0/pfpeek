<#
.SYNOPSIS
    Portfolio performance report for Powershell terminal.

.DESCRIPTION
           __               _
     _ __ / _|_ __  ___ ___| |__
    | '_ \  _| '_ \/ -_) -_) / /
    | .__/_| | .__/\___\___|_\_\
    |_|      |_|

    Retrieves real-time and historical financial data from Yahoo Finance or local CSV exports, calculates portfolio performance metrics, and renders sparkline charts and tables.

    FEATURES:
        * Real-time and historical portfolio tracking.
        * Automatic integration with Directa Trading CSV export.
        * Adaptive console layout.
        * Braille-based sparkline charts for intraday and long-term performance.
        * Comprehensive financial metrics including XIRR/MWRR, expenses, and trades breakdown.
        * Yahoo Finance retries, cached-quote fallback, and visible stale or missing quote status.
        * Automated TSV logging of portfolio metrics and quote quality.

    WHAT'S NEW IN v2:
        * Refreshes periodically.
        * Layout adapts to smaller consoles / bigger fonts.
        * Yahoo requests report clearer failures, retry transient errors with bounded backoff, and reuse cached quotes when available.
        * Stale and missing quotes are identified in the report and recorded in the TSV log.

.NOTES
    AUTHOR:      Raffaele Bianco, GitHub Copilot, Claude, Gemini
    VERSION:     2.00 (2026-11-19)
    BLOG POST:   https://www.raffaelebianco.it/blog/p/pfpeek/
    GITHUB REPO: https://github.com/RaffaeleBianc0/pfpeek
#>

[CmdletBinding()]
param([switch]$CoreOnly)


# === CONFIGURATION STARTS HERE =============================================

# Portfolio source:
#   $true  -> rebuilds it by reading the "Movimenti" CSV export from Directa Trading
#   $false -> uses the static $manualAssets list below
$useCsvPortfolio = $true

# --- Option A: Portfolio from CSV (used if $useCsvPortfolio = $true) ---
# - If $csvPath is specified, that exact file is used.
# - If $csvPath is $null, the script automatically searches for the most recent
#   "Movimenti*.csv" file in the script directory (or current directory).
$csvPath = $null        # e.g. "C:\Path\Movimenti_J1234_20-7-2026.csv"
$tickerSuffix = ".MI"   # Market suffix to add to the tickers read from the CSV (Directa exports them without suffix)

# --- Option A: Manual transactions injected alongside the CSV ones ---
# Use this to add transactions that were never tracked in the Directa CSV export
# (e.g., assets bought on another broker/exchange). Read from a CSV file with the
# SAME columns/format/delimiter (";") as the Directa "Movimenti" export (Italian
# number formatting for Quantita/ImportoEuro/ImportoDivisa, e.g. "1.234,56", and
# dates in "dd-MM-yyyy" format - no header/footer boilerplate, just the column
# header row followed by the transaction rows).
# NOTE: Ticker here must NOT have $tickerSuffix appended (it's added automatically,
# just like for CSV-sourced tickers), unless the asset doesn't need one (e.g. BTC-USD).
#   - If $manualTransactionsCSV is specified, that exact file is used.
#   - If $manualTransactionsCSV is $null/empty, "movimentiAggiuntivi.csv" in the
#     script directory (or current directory) is used; if it doesn't exist yet,
#     it's created automatically with just the header row, may be useful for the future.
$manualTransactionsCSV = ""   # e.g., "C:\Path\movimentiExtra.csv"

# --- Option A: Extra assets added to the ones read from CSV ---
# --- Option B: Manual portfolio ---
$manualAssets = @(
    # @{ Ticker = "XD9U.MI"; AvgCost = 172.55; Qty = 49.00 }
    # @{ Ticker = "EXUS.MI"; AvgCost = 33.27;  Qty = 156.00 }
    # @{ Ticker = "EIMI.MI"; AvgCost = 34.64;  Qty = 43.00 }
    # @{ Ticker = "VWCE.MI"; AvgCost = 163.63; Qty = 1000.00 }
    # @{ Ticker = "XGLE.MI"; AvgCost = 207.00; Qty = 200.00 }
    # @{ Ticker = "WGLD.MI"; AvgCost = 261.57; Qty = 3.00 }
    # @{ Ticker = "XEON.MI"; AvgCost = 144.22; Qty = 35.00 }
    @{ Ticker = "BTC-USD"; AvgCost = 91755.98; Qty = 0.001197 }
)

$assetsSorting = "VWCE.MI", "SWDA.MI", "XD9U.MI", "EXUS.MI", "EIMI.MI", "XGLE.MI", "WGLD.MI", "XEON.MI", "BTC-USD"

# Refresh interval in minutes.
$refreshCountdown = 15
$configuredCsvPath = $csvPath

$refreshIntervalMinutes = 0
if (-not [int]::TryParse([string]$refreshCountdown, [ref]$refreshIntervalMinutes) -or $refreshIntervalMinutes -lt 1) {
    throw "Configuration error: refreshCountdown must be a positive whole number of minutes."
}
$refreshCountdown = $refreshIntervalMinutes
$minWidth = 102

# Custom start date for the overall chart
# (put your first investment date here, or leave empty to use the oldest date read from the CSV):
$startDateConfig = "03/08/2024"

# Logging folder (if $null, the TSV log is saved in the same folder as the script):
$logFolder = $null   # Common alternative: $env:TEMP

# === CONFIGURATION ENDS HERE ===============================================

$script:DirectaCsvHeaderFields = @(
    "DataOperazione", "DataValuta", "TipoOperazione", "Ticker", "Isin",
    "Protocollo", "Descrizione", "Quantita", "ImportoEuro", "ImportoDivisa",
    "Divisa", "RiferimentoOrdine"
)

$script:lastGoodQuotes = @{}

function Get-VisibleLength {
    param([AllowNull()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return 0 }
    return ([regex]::Replace($Text, "\x1b\[[0-9;]*[a-zA-Z]", "")).Length
}

function Get-IntradayTimelineLayout {
    param([ValidateRange(6, 32767)][int]$TargetWidth)

    $points = @(
        [pscustomobject]@{ Label = "Prev"; Hour = 9; Min = 0; IsPrev = $true }
        [pscustomobject]@{ Label = "09"; Hour = 9; Min = 0; IsPrev = $false }
        [pscustomobject]@{ Label = "10"; Hour = 10; Min = 0; IsPrev = $false }
        [pscustomobject]@{ Label = "11"; Hour = 11; Min = 0; IsPrev = $false }
        [pscustomobject]@{ Label = "12"; Hour = 12; Min = 0; IsPrev = $false }
        [pscustomobject]@{ Label = "13"; Hour = 13; Min = 0; IsPrev = $false }
        [pscustomobject]@{ Label = "14"; Hour = 14; Min = 0; IsPrev = $false }
        [pscustomobject]@{ Label = "15"; Hour = 15; Min = 0; IsPrev = $false }
        [pscustomobject]@{ Label = "16"; Hour = 16; Min = 0; IsPrev = $false }
        [pscustomobject]@{ Label = "17"; Hour = 17; Min = 0; IsPrev = $false }
        [pscustomobject]@{ Label = "17:30"; Hour = 17; Min = 30; IsPrev = $false }
    )

    $minuteLabelsFit = $true
    $previousEnd = -1
    $hasPreviousLabel = $false
    foreach ($point in $points) {
        if (-not $point.IsPrev -and $point.Hour -eq 9 -and $point.Min -eq 0) { continue }
        $label = if ($point.IsPrev -or ($point.Hour -eq 17 -and $point.Min -eq 30)) {
            $point.Label
        }
        else {
            "{0:00}:{1:00}" -f $point.Hour, $point.Min
        }

        $percent = if ($point.IsPrev) {
            0.0
        }
        else {
            $totalMinutesStart = 9 * 60
            $totalMinutesEnd = (17 * 60) + 30
            $currentMinutes = $point.Hour * 60 + $point.Min
            0.12 + (0.88 * (($currentMinutes - $totalMinutesStart) / ($totalMinutesEnd - $totalMinutesStart)))
        }
        $consoleX = [int][math]::Round($percent * ($TargetWidth - 1))
        $startPos = [math]::Max(0, $consoleX - [int]($label.Length / 2))
        if ($startPos + $label.Length -gt $TargetWidth) { $startPos = $TargetWidth - $label.Length }
        if ($hasPreviousLabel -and $startPos -le ($previousEnd + 1)) {
            $minuteLabelsFit = $false
            break
        }
        $previousEnd = $startPos + $label.Length - 1
        $hasPreviousLabel = $true
    }

    $labels = [System.Collections.Generic.List[object]]::new()
    foreach ($point in $points) {
        if (-not $point.IsPrev -and $point.Hour -eq 9 -and $point.Min -eq 0) { continue }
        $label = if ($point.IsPrev -or ($point.Hour -eq 17 -and $point.Min -eq 30)) {
            $point.Label
        }
        elseif ($minuteLabelsFit) {
            "{0:00}:{1:00}" -f $point.Hour, $point.Min
        }
        else {
            "{0:00}" -f $point.Hour
        }
        $percent = if ($point.IsPrev) {
            0.0
        }
        else {
            $totalMinutesStart = 9 * 60
            $totalMinutesEnd = (17 * 60) + 30
            $currentMinutes = $point.Hour * 60 + $point.Min
            0.12 + (0.88 * (($currentMinutes - $totalMinutesStart) / ($totalMinutesEnd - $totalMinutesStart)))
        }
        $consoleX = [int][math]::Round($percent * ($TargetWidth - 1))
        $startPos = [math]::Max(0, $consoleX - [int]($label.Length / 2))
        if ($startPos + $label.Length -gt $TargetWidth) { $startPos = $TargetWidth - $label.Length }
        $labels.Add([pscustomobject]@{ Label = $label; Start = $startPos })
    }

    return [pscustomobject]@{
        UsesMinuteLabels = $minuteLabelsFit
        Labels           = $labels.ToArray()
    }
}

function Resolve-QuoteSnapshot {
    param(
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$FreshQuotes,
        [Parameter(Mandatory = $true)][object[]]$Assets,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Cache,
        [datetime]$FetchedAt = (Get-Date)
    )

    $resolvedQuotes = @{}
    $staleQuotes = @{}
    $missingQuotes = [System.Collections.Generic.List[string]]::new()
    foreach ($asset in $Assets) {
        $ticker = [string]$asset.Ticker
        if ($FreshQuotes.Contains($ticker) -and $null -ne $FreshQuotes[$ticker]) {
            $quote = $FreshQuotes[$ticker]
            $Cache[$ticker] = [pscustomobject]@{ Quote = $quote; FetchedAt = $FetchedAt }
            $resolvedQuotes[$ticker] = $quote
        }
        elseif ($Cache.Contains($ticker) -and $null -ne $Cache[$ticker].Quote) {
            $cachedQuote = $Cache[$ticker]
            $resolvedQuotes[$ticker] = $cachedQuote.Quote
            $age = $FetchedAt - [datetime]$cachedQuote.FetchedAt
            $staleQuotes[$ticker] = [pscustomobject]@{
                Quote      = $cachedQuote.Quote
                FetchedAt  = $cachedQuote.FetchedAt
                AgeSeconds = [math]::Max(0, $age.TotalSeconds)
            }
        }
        else {
            $missingQuotes.Add($ticker)
        }
    }

    return [pscustomobject]@{
        Quotes  = $resolvedQuotes
        Stale   = $staleQuotes
        Missing = $missingQuotes.ToArray()
    }
}

function Get-YahooResponseStatusCode {
    param([Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord]$ErrorRecord)
    $response = $ErrorRecord.Exception.Response
    if ($null -eq $response) { return $null }
    $statusCode = $response.PSObject.Properties["StatusCode"]
    if ($null -eq $statusCode) { return $null }
    return [int]$statusCode.Value
}

function Get-YahooRequestFailureSummary {
    param([Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord]$ErrorRecord)
    $statusCode = Get-YahooResponseStatusCode -ErrorRecord $ErrorRecord
    if ($statusCode -eq 429) { return "HTTP 429 rate limit" }
    if ($statusCode -eq 401 -or $statusCode -eq 403) { return "HTTP $statusCode authorization/consent failure" }
    if ($null -ne $statusCode) { return "HTTP $statusCode" }
    return $ErrorRecord.Exception.Message
}

function Invoke-YahooRequestWithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Request,
        [ValidateRange(1, 5)][int]$MaxAttempts = 3
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return & $Request
        }
        catch {
            $statusCode = Get-YahooResponseStatusCode -ErrorRecord $_
            $isTransient = $null -eq $statusCode -or $statusCode -eq 408 -or $statusCode -eq 429 -or $statusCode -ge 500
            if (-not $isTransient -or $attempt -eq $MaxAttempts) { throw }

            $delaySeconds = [int][math]::Min(8, [math]::Pow(2, $attempt - 1))
            if ($statusCode -eq 429) {
                $response = $_.Exception.Response
                $headers = if ($null -ne $response) { $response.PSObject.Properties["Headers"] } else { $null }
                $retryAfter = if ($null -ne $headers -and $null -ne $headers.Value) { $headers.Value["Retry-After"] } else { $null }
                $retryAfterSeconds = 0
                if ([int]::TryParse([string]$retryAfter, [ref]$retryAfterSeconds)) {
                    $delaySeconds = [int][math]::Min(8, [math]::Max(1, $retryAfterSeconds))
                }
                else {
                    $retryAfterDate = [datetime]::MinValue
                    if ([datetime]::TryParse([string]$retryAfter, [ref]$retryAfterDate)) {
                        $delaySeconds = [int][math]::Min(8, [math]::Max(1, [math]::Ceiling(($retryAfterDate - (Get-Date)).TotalSeconds)))
                    }
                }
            }
            Start-Sleep -Seconds $delaySeconds
        }
    }
}

function Get-SegmentLayout {
    param(
        [string[]]$Segments,
        [int]$ConsoleWidth = [int]::MaxValue,
        [string]$Separator = "  ",
        [string]$ContinuationIndent = "  ",
        [bool[]]$BreaksBefore = $null
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $breaks = if ($null -ne $BreaksBefore) { [bool[]]$BreaksBefore.Clone() } else { New-Object bool[] ($Segments.Count) }
    $currentLine = $null
    $currentLength = 0
    $separatorLength = Get-VisibleLength $Separator
    $indentLength = Get-VisibleLength $ContinuationIndent

    for ($i = 0; $i -lt $Segments.Count; $i++) {
        $segment = $Segments[$i]
        if ([string]::IsNullOrEmpty($segment)) { continue }
        $segmentLength = Get-VisibleLength $segment

        if ($null -eq $currentLine) {
            $currentLine = $segment
            $currentLength = $segmentLength
            continue
        }

        $wrapBefore = if ($null -ne $BreaksBefore) {
            $i -lt $breaks.Count -and $breaks[$i]
        }
        else {
            ($currentLength + $separatorLength + $segmentLength) -gt $ConsoleWidth
        }

        if ($wrapBefore) {
            if ($null -eq $BreaksBefore) { $breaks[$i] = $true }
            $lines.Add($currentLine)
            $currentLine = $ContinuationIndent + $segment
            $currentLength = $indentLength + $segmentLength
        }
        else {
            $currentLine += $Separator + $segment
            $currentLength += $separatorLength + $segmentLength
        }
    }

    if ($null -ne $currentLine) { $lines.Add($currentLine) }
    return [pscustomobject]@{
        Lines        = $lines.ToArray()
        BreaksBefore = $breaks
    }
}

function Get-WrappedSegmentLines {
    param(
        [string[]]$Segments,
        [int]$ConsoleWidth,
        [string]$Separator = "  ",
        [string]$ContinuationIndent = "  "
    )
    $layout = Get-SegmentLayout -Segments $Segments -ConsoleWidth $ConsoleWidth -Separator $Separator -ContinuationIndent $ContinuationIndent
    return , $layout.Lines
}

function Get-SegmentBreaksBefore {
    param(
        [string[]]$Segments,
        [int]$ConsoleWidth,
        [string]$Separator = "  ",
        [string]$ContinuationIndent = "  "
    )
    $layout = Get-SegmentLayout -Segments $Segments -ConsoleWidth $ConsoleWidth -Separator $Separator -ContinuationIndent $ContinuationIndent
    return , $layout.BreaksBefore
}

function Get-LinesFromBreaks {
    param(
        [string[]]$Segments,
        [bool[]]$BreaksBefore,
        [string]$Separator = "  ",
        [string]$ContinuationIndent = "  "
    )
    $layout = Get-SegmentLayout -Segments $Segments -BreaksBefore $BreaksBefore -Separator $Separator -ContinuationIndent $ContinuationIndent
    return , $layout.Lines
}

function ConvertTo-ItalianDouble {
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return 0.0 }
    $clean = $Text.Trim().Replace(".", "").Replace(",", ".")
    $value = 0.0
    if (-not [double]::TryParse(
            $clean,
            [System.Globalization.NumberStyles]::Any,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [ref]$value
        )) {
        throw "Invalid Italian-formatted number: '$Text'."
    }
    return $value
}

function ConvertFrom-DirectaCsvLines {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Lines,
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [switch]$NormalizeDateSeparators
    )

    $headerIndex = -1
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Lines[$i].TrimStart([char]0xFEFF) -match '^Data operazione;') {
            $headerIndex = $i
            break
        }
    }
    if ($headerIndex -lt 0) {
        throw "Header 'Data operazione' not found in CSV: $Path"
    }

    $dataLines = @()
    if ($headerIndex + 1 -lt $Lines.Count) {
        $dataLines = @($Lines[($headerIndex + 1)..($Lines.Count - 1)] |
            Where-Object { $_.Trim().Length -gt 0 })
    }
    if ($dataLines.Count -eq 0) { return , @() }

    $transactions = @($dataLines | ConvertFrom-Csv -Delimiter ";" -Header $script:DirectaCsvHeaderFields)
    if ($NormalizeDateSeparators) {
        foreach ($transaction in $transactions) {
            $transaction.DataOperazione = $transaction.DataOperazione -replace '/', '-'
            $transaction.DataValuta = $transaction.DataValuta -replace '/', '-'
        }
    }
    return , $transactions
}

function Resolve-DirectaCsvPath {
    param(
        [string]$ConfiguredPath,
        [Parameter(Mandatory = $true)]
        [string]$SearchDirectory
    )

    if (-not [string]::IsNullOrWhiteSpace($ConfiguredPath)) { return $ConfiguredPath }
    $latestCsv = Get-ChildItem -Path $SearchDirectory -Filter "Movimenti*.csv" -File -ErrorAction Stop |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
    if ($null -eq $latestCsv) {
        throw "No 'Movimenti*.csv' file found in '$SearchDirectory' and no configured CSV path was provided."
    }
    return $latestCsv.FullName
}

function Get-TransactionSummary {
    param([Parameter(Mandatory = $true)][object[]]$Transactions)

    $summary = [ordered]@{
        BuyCount         = 0
        SellCount        = 0
        Commissions      = 0.0
        StampDuty        = 0.0
        WithholdingTaxes = 0.0
        Deposited        = 0.0
        Withdrawn        = 0.0
    }
    foreach ($transaction in $Transactions) {
        switch -Wildcard ($transaction.TipoOperazione) {
            "Acquisto" { $summary.BuyCount++ }
            "Vendita" { $summary.SellCount++ }
            "Commissioni" { $summary.Commissions += ConvertTo-ItalianDouble $transaction.ImportoEuro }
            "Bollo portafoglio titoli*" { $summary.StampDuty += ConvertTo-ItalianDouble $transaction.ImportoEuro }
            "Rit. etf" { $summary.WithholdingTaxes += ConvertTo-ItalianDouble $transaction.ImportoEuro }
            "Conferimento con bonifico" { $summary.Deposited += ConvertTo-ItalianDouble $transaction.ImportoEuro }
            "Prelievo bonifico" { $summary.Withdrawn += ConvertTo-ItalianDouble $transaction.ImportoEuro }
        }
    }
    return [pscustomobject]$summary
}

function Get-CostRatio {
    param(
        [double]$TotalGainAllTime,
        [double]$TotalRealizedGain,
        [double]$TotalExpenses,
        [double]$TotalDeposited,
        [double]$TotalCost
    )

    if ($TotalDeposited -ne 0) {
        if ($TotalGainAllTime -le 0 -or $TotalRealizedGain -le 0) { return $null }
        return ($TotalExpenses / $TotalRealizedGain) * 100
    }
    if ($TotalCost -ne 0) { return ($TotalExpenses / $TotalCost) * 100 }
    return $null
}

function ConvertTo-PortfolioSideFlows {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$InvestorFlows)

    $byDate = @{}
    foreach ($flow in $InvestorFlows) {
        $date = $flow.Date.Date
        $key = $date.ToString("yyyyMMdd", [System.Globalization.CultureInfo]::InvariantCulture)
        if (-not $byDate.ContainsKey($key)) {
            $byDate[$key] = [pscustomobject]@{ Date = $date; Amount = 0.0 }
        }
        $portfolioAmount = if ($flow.Type -eq "Cost") { $flow.Amount } else { - $flow.Amount }
        $byDate[$key].Amount += $portfolioAmount
    }
    return , @($byDate.Values | Sort-Object Date)
}

function Add-PortfolioLogEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Stats,
        [Parameter(Mandatory = $true)]
        [string]$Timestamp
    )

    $expectedHeader = @("Timestamp", "LogLevel") + @($Stats.Keys)
    if (-not (Test-Path -LiteralPath $Path)) {
        Set-Content -LiteralPath $Path -Value ($expectedHeader -join "`t") -Encoding UTF8 -ErrorAction Stop
    }
    else {
        $existingLines = @(Get-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop)
        if ($existingLines.Count -eq 0) {
            Set-Content -LiteralPath $Path -Value ($expectedHeader -join "`t") -Encoding UTF8 -ErrorAction Stop
        }
        else {
            $existingHeader = $existingLines[0].Split([char]9)
            if (($existingHeader -join "`t") -ne ($expectedHeader -join "`t")) {
                if ($existingHeader.Count -gt $expectedHeader.Count) {
                    throw "Portfolio log schema has more columns than the current writer supports."
                }
                for ($columnIndex = 0; $columnIndex -lt $existingHeader.Count; $columnIndex++) {
                    if ($existingHeader[$columnIndex] -ne $expectedHeader[$columnIndex]) {
                        throw "Portfolio log schema mismatch at column $($columnIndex + 1): expected '$($expectedHeader[$columnIndex])', found '$($existingHeader[$columnIndex])'."
                    }
                }

                $updatedLines = [System.Collections.Generic.List[string]]::new()
                $updatedLines.Add(($expectedHeader -join "`t"))
                for ($lineIndex = 1; $lineIndex -lt $existingLines.Count; $lineIndex++) {
                    $fields = [System.Collections.Generic.List[string]]::new()
                    $fields.AddRange([string[]]$existingLines[$lineIndex].Split([char]9))
                    if ($fields.Count -gt $expectedHeader.Count) {
                        throw "Portfolio log row $($lineIndex + 1) has more fields than the current header."
                    }
                    while ($fields.Count -lt $expectedHeader.Count) { $fields.Add("") }
                    $updatedLines.Add(($fields -join "`t"))
                }

                $directory = Split-Path -Parent $Path
                if ([string]::IsNullOrWhiteSpace($directory)) { $directory = (Get-Location).Path }
                $migrationPath = Join-Path $directory ([System.IO.Path]::GetRandomFileName())
                try {
                    Set-Content -LiteralPath $migrationPath -Value $updatedLines.ToArray() -Encoding UTF8 -ErrorAction Stop
                    Move-Item -LiteralPath $migrationPath -Destination $Path -Force -ErrorAction Stop
                }
                finally {
                    if (Test-Path -LiteralPath $migrationPath) {
                        Remove-Item -LiteralPath $migrationPath -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    }
    $line = (@($Timestamp, "INFO") + @($Stats.Values)) -join "`t"
    Add-Content -LiteralPath $Path -Value $line -Encoding UTF8 -ErrorAction Stop
}

function Find-NearestSortedIndex {
    param(
        [Parameter(Mandatory = $true)]
        [long[]]$Timestamps,
        [Parameter(Mandatory = $true)]
        [double]$Target
    )

    if ($Timestamps.Count -eq 0) { return -1 }

    $low = 0
    $high = $Timestamps.Count - 1
    while ($low -le $high) {
        $mid = $low + [int][math]::Floor(($high - $low) / 2)
        if ($Timestamps[$mid] -eq $Target) { return $mid }
        if ($Timestamps[$mid] -lt $Target) {
            $low = $mid + 1
        }
        else {
            $high = $mid - 1
        }
    }

    if ($high -lt 0) { return 0 }
    if ($low -ge $Timestamps.Count) { return $Timestamps.Count - 1 }
    $beforeDiff = [math]::Abs([double]$Timestamps[$high] - $Target)
    $afterDiff = [math]::Abs([double]$Timestamps[$low] - $Target)
    if ($beforeDiff -le $afterDiff) { return $high }
    return $low
}

function Find-LastCheckpointIndex {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Checkpoints,
        [Parameter(Mandatory = $true)]
        [long]$Target
    )

    $low = 0
    $high = $Checkpoints.Count - 1
    $result = -1
    while ($low -le $high) {
        $mid = $low + [int][math]::Floor(($high - $low) / 2)
        if ([long]$Checkpoints[$mid].DateUnix -le $Target) {
            $result = $mid
            $low = $mid + 1
        }
        else {
            $high = $mid - 1
        }
    }
    return $result
}

function Get-AlignedHistoricalPrices {
    param(
        [Parameter(Mandatory = $true)]
        [long[]]$TargetTimestamps,
        [Parameter(Mandatory = $true)]
        [long[]]$SourceTimestamps,
        [Parameter(Mandatory = $true)]
        [double[]]$SourcePrices,
        [double]$FallbackPrice = 0.0
    )

    $aligned = [double[]]::new($TargetTimestamps.Count)
    $sourceCount = [math]::Min($SourceTimestamps.Count, $SourcePrices.Count)
    $sourceIndex = 0
    $lastPrice = $FallbackPrice

    for ($targetIndex = 0; $targetIndex -lt $TargetTimestamps.Count; $targetIndex++) {
        while ($sourceIndex -lt $sourceCount -and $SourceTimestamps[$sourceIndex] -le $TargetTimestamps[$targetIndex]) {
            $lastPrice = $SourcePrices[$sourceIndex]
            $sourceIndex++
        }
        $aligned[$targetIndex] = $lastPrice
    }
    return , $aligned
}

function New-IndependentWebRequestSession {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$SourceSession,
        [Parameter(Mandatory = $true)]
        [uri]$CookieUri
    )

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $userAgent = $SourceSession.Headers["User-Agent"]
    if ($userAgent) { $session.Headers.Add("User-Agent", $userAgent) }
    foreach ($cookie in $SourceSession.Cookies.GetCookies($CookieUri)) {
        $session.Cookies.Add($CookieUri, $cookie)
    }
    return $session
}

function Start-RunspaceScript {
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.RunspacePool]$RunspacePool,
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        [Parameter(Mandatory = $true)]
        [hashtable]$Parameters
    )

    $pipeline = [powershell]::Create()
    try {
        $pipeline.RunspacePool = $RunspacePool
        [void]$pipeline.AddScript($ScriptBlock).AddParameters($Parameters)
        $handle = $pipeline.BeginInvoke()
        return [pscustomobject]@{ Pipe = $pipeline; Handle = $handle }
    }
    catch {
        $pipeline.Dispose()
        throw
    }
}

function Get-XirrRobust {
    param([Parameter(Mandatory = $true)][object[]]$Cashflows)

    if ($Cashflows.Count -lt 2) { return $null }
    $sorted = @($Cashflows | Sort-Object Date)
    $hasPositive = @($sorted | Where-Object { $_.Amount -gt 0 }).Count -gt 0
    $hasNegative = @($sorted | Where-Object { $_.Amount -lt 0 }).Count -gt 0
    if (-not $hasPositive -or -not $hasNegative) { return $null }

    $baseDate = $sorted[0].Date
    $amountScale = [math]::Max(1.0, ($sorted | ForEach-Object { [math]::Abs([double]$_.Amount) } | Measure-Object -Sum).Sum)
    $npvTolerance = $amountScale * 1e-10

    function Get-XirrNpv([double]$Rate) {
        if ($Rate -le -1.0) { return [double]::NaN }
        $sum = 0.0
        foreach ($cashflow in $sorted) {
            $days = ($cashflow.Date - $baseDate).TotalDays
            $sum += $cashflow.Amount / [math]::Pow((1.0 + $Rate), ($days / 365.0))
        }
        return $sum
    }

    $guesses = @(-0.99, -0.5, -0.1, 0.0, 0.01, 0.05, 0.1, 0.15, 0.2, 0.3, 0.5, 1.0, 2.0, 5.0)
    foreach ($guess in $guesses) {
        $rate = [double]$guess
        for ($iteration = 0; $iteration -lt 200; $iteration++) {
            $npv = Get-XirrNpv $rate
            if ([double]::IsNaN($npv) -or [double]::IsInfinity($npv)) { break }
            $epsilon = 0.00001
            $npvEpsilon = Get-XirrNpv ($rate + $epsilon)
            $derivative = ($npvEpsilon - $npv) / $epsilon
            if ([double]::IsNaN($derivative) -or [math]::Abs($derivative) -lt 1e-12) { break }

            $newRate = $rate - $npv / $derivative
            if ([double]::IsNaN($newRate) -or [double]::IsInfinity($newRate) -or $newRate -le -0.999 -or $newRate -ge 100) { break }
            if ([math]::Abs($newRate - $rate) -lt 1e-10) {
                if ([math]::Abs((Get-XirrNpv $newRate)) -le $npvTolerance) { return $newRate }
                break
            }
            $rate = $newRate
        }
    }

    $low = -0.99
    $high = 10.0
    $npvLow = Get-XirrNpv $low
    $npvHigh = Get-XirrNpv $high
    if ([double]::IsNaN($npvLow) -or [double]::IsNaN($npvHigh) -or $npvLow * $npvHigh -ge 0) { return $null }

    for ($iteration = 0; $iteration -lt 100; $iteration++) {
        $mid = ($low + $high) / 2.0
        $npvMid = Get-XirrNpv $mid
        if ([double]::IsNaN($npvMid) -or [double]::IsInfinity($npvMid)) { return $null }
        if ([math]::Abs($npvMid) -le $npvTolerance) { return $mid }

        if ($npvLow * $npvMid -lt 0) {
            $high = $mid
            $npvHigh = $npvMid
        }
        else {
            $low = $mid
            $npvLow = $npvMid
        }
    }
    return $null
}

function Get-QtyHistory {
    param([object[]]$Transactions)

    $parsed = [System.Collections.Generic.List[PSCustomObject]]::new()
    $transactionIndex = 0
    foreach ($transaction in $Transactions) {
        if ($transaction.TipoOperazione -ne "Acquisto" -and $transaction.TipoOperazione -ne "Vendita") { continue }
        if ([string]::IsNullOrWhiteSpace($transaction.Ticker)) { continue }
        $date = [datetime]::MinValue
        if (-not [datetime]::TryParseExact($transaction.DataOperazione, "dd-MM-yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$date)) { continue }
        $quantity = ConvertTo-ItalianDouble $transaction.Quantita
        $delta = if ($transaction.TipoOperazione -eq "Acquisto") { $quantity } else { - $quantity }
        $parsed.Add([pscustomobject]@{ Ticker = $transaction.Ticker; Date = $date; Delta = $delta; Index = $transactionIndex })
        $transactionIndex++
    }

    $history = @{}
    $running = @{}
    foreach ($transaction in ($parsed | Sort-Object Date, Index)) {
        if (-not $running.ContainsKey($transaction.Ticker)) { $running[$transaction.Ticker] = 0.0 }
        $running[$transaction.Ticker] += $transaction.Delta
        if (-not $history.ContainsKey($transaction.Ticker)) {
            $history[$transaction.Ticker] = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
        $timestamp = (New-Object DateTimeOffset ($transaction.Date)).ToUnixTimeSeconds()
        $history[$transaction.Ticker].Add([pscustomobject]@{ DateUnix = $timestamp; Qty = $running[$transaction.Ticker] })
    }
    return $history
}

function Get-CashHistory {
    param([object[]]$Transactions)

    $history = [System.Collections.Generic.List[PSCustomObject]]::new()
    $index = 0
    foreach ($transaction in $Transactions) {
        $date = [datetime]::MinValue
        if (-not [datetime]::TryParseExact($transaction.DataOperazione, "dd-MM-yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$date)) { continue }
        $amount = ConvertTo-ItalianDouble $transaction.ImportoEuro
        switch -Wildcard ($transaction.TipoOperazione) {
            "Conferimento con bonifico" { $delta = [math]::Abs($amount); $type = "External" }
            "Prelievo bonifico" { $delta = - [math]::Abs($amount); $type = "External" }
            "Commissioni" { $delta = - [math]::Abs($amount); $type = "Cost" }
            "Bollo portafoglio titoli*" { $delta = - [math]::Abs($amount); $type = "Cost" }
            "Rit. etf" { $delta = - [math]::Abs($amount); $type = "Cost" }
            "Acquisto" { $delta = - [math]::Abs($amount); $type = "Internal" }
            "Vendita" { $delta = [math]::Abs($amount); $type = "Internal" }
            default { continue }
        }
        $timestamp = (New-Object DateTimeOffset ($date)).ToUnixTimeSeconds()
        $history.Add([pscustomobject]@{ DateUnix = $timestamp; Delta = $delta; Type = $type; Index = $index })
        $index++
    }

    $sorted = @($history | Sort-Object DateUnix, Index)
    if ($sorted.Count -gt 0) {
        $firstTimestamp = $sorted[0].DateUnix
        $sameDayPoints = @($sorted | Where-Object { $_.DateUnix -eq $firstTimestamp })
        $hasExternalOnFirstDate = @($sameDayPoints | Where-Object { $_.Type -eq "External" })
        if ($hasExternalOnFirstDate.Count -eq 0) {
            $netFirstDayDelta = ($sameDayPoints | Measure-Object -Property Delta -Sum).Sum
            if ($netFirstDayDelta -lt 0) {
                $sorted = @([pscustomobject]@{ DateUnix = $firstTimestamp; Delta = - $netFirstDayDelta; Type = "External"; Index = -1 }) + $sorted
            }
        }
    }

    $runningBalance = 0.0
    foreach ($point in $sorted) {
        $runningBalance += $point.Delta
        $point | Add-Member -MemberType NoteProperty -Name Balance -Value $runningBalance -Force
    }
    return , [System.Collections.Generic.List[PSCustomObject]]$sorted
}

function Get-CashBalanceAtUnix {
    param($TargetUnix, $CashHistory)
    if (-not $CashHistory -or $CashHistory.Count -eq 0) { return 0.0 }
    $index = Find-LastCheckpointIndex -Checkpoints $CashHistory -Target $TargetUnix
    if ($index -lt 0) { return 0.0 }
    return $CashHistory[$index].Balance
}

function Get-QtyAtUnix {
    param($TickerHistory, $TargetUnix)
    if (-not $TickerHistory -or $TickerHistory.Count -eq 0) { return 0.0 }
    $index = Find-LastCheckpointIndex -Checkpoints $TickerHistory -Target $TargetUnix
    if ($index -lt 0) { return 0.0 }
    return $TickerHistory[$index].Qty
}

function Get-HistoricalPortfolioValue {
    param($TargetUnix, $QtyHistory, $PriceSeriesByTicker, $TickerSuffix, $CashHistory)

    $total = 0.0
    foreach ($fullTicker in $PriceSeriesByTicker.Keys) {
        $rawTicker = if ($fullTicker.EndsWith($TickerSuffix)) {
            $fullTicker.Substring(0, $fullTicker.Length - $TickerSuffix.Length)
        }
        else {
            $fullTicker
        }
        $quantity = Get-QtyAtUnix $QtyHistory[$rawTicker] $TargetUnix
        if ($quantity -le 0) { continue }
        $series = $PriceSeriesByTicker[$fullTicker]
        $index = Find-NearestSortedIndex -Timestamps $series.Timestamps -Target $TargetUnix
        if ($index -lt 0 -or $index -ge $series.Closes.Count) { continue }
        $total += [double]$series.Closes[$index] * $quantity
    }
    return ($total + (Get-CashBalanceAtUnix $TargetUnix $CashHistory))
}

function Get-TwrrRobust {
    param($ExternalFlows, $QtyHistory, $PriceSeriesByTicker, $TickerSuffix, $WindowStartUnix, $EndValue, $EndDate, $CashHistory)

    if (-not $PriceSeriesByTicker -or $PriceSeriesByTicker.Keys.Count -eq 0) {
        Write-Warning "TWRR: skipped - no historical price series available for any traded ticker. Likely a failed/rate-limited Yahoo Finance download."
        return $null
    }

    if ($CashHistory -and $CashHistory.Count -gt 0 -and $CashHistory[0].DateUnix -gt $WindowStartUnix) {
        $WindowStartUnix = $CashHistory[0].DateUnix
    }
    $periodStartValue = Get-HistoricalPortfolioValue $WindowStartUnix $QtyHistory $PriceSeriesByTicker $TickerSuffix $CashHistory
    if ($periodStartValue -le 0) {
        Write-Warning "TWRR: skipped - reconstructed initial portfolio value is zero or negative ($periodStartValue)."
        return $null
    }
    $periodStartDate = ([DateTimeOffset]::FromUnixTimeSeconds([int64]$WindowStartUnix)).LocalDateTime.Date
    $cumulativeFactor = 1.0
    $sortedFlows = ConvertTo-PortfolioSideFlows -InvestorFlows $ExternalFlows

    foreach ($flow in $sortedFlows) {
        $flowUnix = (New-Object DateTimeOffset ($flow.Date)).ToUnixTimeSeconds()
        if ($flowUnix -lt $WindowStartUnix) { continue }
        $valueAtFlow = Get-HistoricalPortfolioValue $flowUnix $QtyHistory $PriceSeriesByTicker $TickerSuffix $CashHistory
        $valueBeforeFlow = $valueAtFlow - $flow.Amount
        $ratio = if ($periodStartValue -ne 0) { $valueBeforeFlow / $periodStartValue } else { [double]::NaN }
        $cumulativeFactor *= $ratio
        $periodStartValue = $valueAtFlow
        if ($periodStartValue -le 0) {
            Write-Warning "TWRR: skipped - sub-period value base hit zero/negative ($periodStartValue) after cash flows on $($flow.Date.ToString('dd/MM/yyyy')). Can't chain returns past a wipeout."
            return $null
        }
    }

    if ($periodStartValue -le 0) {
        Write-Warning "TWRR: skipped - final sub-period value base is zero or negative ($periodStartValue)."
        return $null
    }
    $cumulativeFactor *= ($EndValue / $periodStartValue)
    if ($cumulativeFactor -le 0) {
        Write-Warning "TWRR: skipped - cumulative return factor is zero or negative ($cumulativeFactor); current total value / last value base gave an invalid ratio."
        return $null
    }
    $totalDays = ($EndDate - $periodStartDate).TotalDays
    if ($totalDays -le 0) { return ($cumulativeFactor - 1.0) }
    return ([math]::Pow($cumulativeFactor, 365.0 / $totalDays) - 1.0)
}

function Set-RefreshCountdownDisplay {
    param(
        [datetime]$Deadline,
        [switch]$Clear
    )

    $title = if (-not [string]::IsNullOrEmpty($script:renderedTitle)) {
        $script:renderedTitle
    }
    else {
        "Portfolio performance $((Get-Date).ToString('dd/MM/yyyy HH:mm'))"
    }
    $text = ""
    if (-not $Clear) {
        $remaining = $Deadline - (Get-Date)
        if ($remaining.TotalSeconds -lt 0) { $remaining = [TimeSpan]::Zero }
        $minutes = [int][math]::Floor($remaining.TotalMinutes)
        $seconds = $remaining.Seconds
        $text = "refreshing data in {0:00}m{1:00}s" -f $minutes, $seconds
    }

    try {
        $rawUi = $Host.UI.RawUI
        $windowSize = $rawUi.WindowSize
        $windowPosition = $rawUi.WindowPosition
        $cursorPosition = $rawUi.CursorPosition
        $status = if ($Clear) { "" } else { " ($text)" }
        $lineLength = $title.Length + $status.Length
        $paddingLength = [math]::Max(0, $windowSize.Width - $lineLength)
        $x = $windowPosition.X
        $y = $windowPosition.Y
        $rawUi.CursorPosition = New-Object System.Management.Automation.Host.Coordinates($x, $y)
        Write-Host $title -ForegroundColor White -NoNewline
        if (-not $Clear) { Write-Host $status -ForegroundColor DarkGray -NoNewline }
        Write-Host (" " * $paddingLength) -NoNewline
        $rawUi.CursorPosition = $cursorPosition
    }
    catch {
        if (-not $script:refreshProgressFallbackWarned) {
            Write-Warning "This PowerShell host cannot position console text; showing the refresh countdown as a progress overlay instead."
            $script:refreshProgressFallbackWarned = $true
        }
        if ($Clear) {
            Write-Progress -Id 2 -Activity "Waiting for next refresh" -Completed
        }
        else {
            $totalSeconds = $refreshCountdown * 60
            $remainingSeconds = [math]::Max(0, [int][math]::Ceiling(($Deadline - (Get-Date)).TotalSeconds))
            $percent = [int](100 * ($totalSeconds - $remainingSeconds) / $totalSeconds)
            Write-Progress -Id 2 -Activity "Waiting for next refresh" -Status $text -PercentComplete $percent
        }
    }
}

function Get-ConsoleDimensions {
    try {
        $size = $Host.UI.RawUI.WindowSize
        if ($size.Width -gt 0 -and $size.Height -gt 0) {
            return [pscustomobject]@{ Width = $size.Width; Height = $size.Height }
        }
    }
    catch {
        if (-not $script:resizeDetectionWarningShown) {
            Write-Warning "This PowerShell host cannot detect console resizing; automatic redraw is unavailable."
            $script:resizeDetectionWarningShown = $true
        }
    }
    return $null
}

function Wait-ForRefresh {
    param(
        [datetime]$Deadline,
        [scriptblock]$OnResize,
        [scriptblock]$DimensionProvider = { Get-ConsoleDimensions }
    )

    $lastObservedSize = if ($script:lastRenderedConsoleWidth -gt 0 -and $script:lastRenderedConsoleHeight -gt 0) {
        [pscustomobject]@{ Width = $script:lastRenderedConsoleWidth; Height = $script:lastRenderedConsoleHeight }
    }
    else {
        & $DimensionProvider
    }
    while ((Get-Date) -lt $Deadline) {
        $currentSize = & $DimensionProvider
        if ($null -ne $currentSize -and (
                $null -eq $lastObservedSize -or
                $currentSize.Width -ne $lastObservedSize.Width -or
                $currentSize.Height -ne $lastObservedSize.Height
            )) {
            if ($null -ne $OnResize) {
                Clear-Host
                if ($currentSize.Width -lt $minWidth) {
                    Show-ConsoleSizeWarning -RequiredWidth $minWidth -RequiredHeight $currentSize.Height
                }
                $currentSize = & $DimensionProvider
                if ($null -eq $currentSize) {
                    throw "Unable to determine console dimensions after resizing."
                }
                if ($currentSize.Width -ge $minWidth) {
                    $script:isResizeRedraw = $true
                    try {
                        & $OnResize
                    }
                    finally {
                        $script:isResizeRedraw = $false
                    }
                }
            }
            $lastObservedSize = & $DimensionProvider
        }
        Set-RefreshCountdownDisplay -Deadline $Deadline
        Start-Sleep -Seconds 1
    }
    Set-RefreshCountdownDisplay -Clear
}

function Show-ConsoleSizeWarning {
    param(
        [int]$RequiredWidth,
        [int]$RequiredHeight
    )

    while ($true) {
        $size = Get-ConsoleDimensions
        if ($null -eq $size) {
            throw "Unable to determine console dimensions while waiting for a resize."
        }
        if ($size.Width -ge $RequiredWidth -and $size.Height -ge $RequiredHeight) {
            Clear-Host
            return
        }

        Clear-Host
        $rawUi = $Host.UI.RawUI
        $windowSize = $rawUi.WindowSize
        $windowPosition = $rawUi.WindowPosition
        $minimumLabel = if ($windowSize.Width -lt 28) { "Min: " } else { "Minimum: " }
        $currentLabel = if ($windowSize.Width -lt 28) { "Now: " } else { "Current: " }
        $rowCount = [math]::Min(3, $windowSize.Height)
        $firstRow = [int][math]::Floor(($windowSize.Height - $rowCount) / 2)
        $drawDimensionLine = {
            param([int]$Row, [string]$Label, [int]$Width, [int]$Height, [ConsoleColor]$WidthColor, [ConsoleColor]$HeightColor)
            $widthText = [string]$Width
            $heightText = [string]$Height
            $lineLength = $Label.Length + $widthText.Length + 3 + $heightText.Length
            $x = $windowPosition.X + [math]::Max(0, [int][math]::Floor(($windowSize.Width - $lineLength) / 2))
            $rawUi.CursorPosition = New-Object System.Management.Automation.Host.Coordinates($x, ($windowPosition.Y + $firstRow + $Row))
            Write-Host $Label -ForegroundColor DarkGray -NoNewline
            Write-Host $widthText -ForegroundColor $WidthColor -NoNewline
            Write-Host " x " -ForegroundColor DarkGray -NoNewline
            Write-Host $heightText -ForegroundColor $HeightColor -NoNewline
        }

        & $drawDimensionLine 0 $minimumLabel $RequiredWidth $RequiredHeight White White
        if ($rowCount -ge 2) {
            $currentWidthColor = if ($size.Width -lt $RequiredWidth) { [ConsoleColor]::Red } else { [ConsoleColor]::Green }
            $currentHeightColor = if ($size.Height -lt $RequiredHeight) { [ConsoleColor]::Red } else { [ConsoleColor]::Green }
            & $drawDimensionLine 1 $currentLabel $size.Width $size.Height $currentWidthColor $currentHeightColor
        }
        if ($rowCount -ge 3) {
            $prompt = "Resize the console to continue"
            $promptX = $windowPosition.X + [math]::Max(0, [int][math]::Floor(($windowSize.Width - $prompt.Length) / 2))
            $rawUi.CursorPosition = New-Object System.Management.Automation.Host.Coordinates($promptX, ($windowPosition.Y + $firstRow + 2))
            Write-Host $prompt -ForegroundColor DarkGray -NoNewline
        }
        Start-Sleep -Seconds 1
    }
}

if ($CoreOnly) { return }

# --- MINIMUM CONSOLE SIZE CHECK ---
$consoleWidth = $minWidth
try {
    if ($Host.UI.RawUI.WindowSize.Width) { $consoleWidth = $Host.UI.RawUI.WindowSize.Width }
}
catch {
    throw "Configuration error: pfpeek requires an interactive PowerShell host with console window dimensions. $_"
}
if ($consoleWidth -lt $minWidth) {
    $startupSize = Get-ConsoleDimensions
    if ($null -eq $startupSize) { throw "Unable to determine console dimensions while waiting for a resize." }
    Show-ConsoleSizeWarning -RequiredWidth $minWidth -RequiredHeight $startupSize.Height
    $startupSize = Get-ConsoleDimensions
    if ($null -eq $startupSize) { throw "Unable to determine console dimensions after resizing." }
    $consoleWidth = $startupSize.Width
}

# --- LOG FOLDER WRITE ACCESS CHECK ---
$targetLogDir = if ($null -ne $logFolder -and $logFolder.Trim() -ne "") {
    $logFolder
}
else {
    if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
}

if (-not (Test-Path $targetLogDir)) {
    try {
        [void][System.IO.Directory]::CreateDirectory($targetLogDir)
    }
    catch {
        Write-Error "ERROR: Cannot create log directory '$targetLogDir': $_"
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }
}

try {
    $testFile = Join-Path $targetLogDir ([System.IO.Path]::GetRandomFileName())
    [System.IO.File]::WriteAllText($testFile, "test")
    [System.IO.File]::Delete($testFile)
}
catch {
    Write-Error "ERROR: No write access to the log folder '$targetLogDir': $_"
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# --- PROGRESS TRACKING ---
function Show-Progress {
    param (
        [int]$Current,
        [int]$Total,
        [string]$Activity = "Processing",
        [string]$Highlight = $null
    )

    if ($Total -le 0) { return }

    $percentComplete = [int][math]::Min(100, [math]::Max(0, ($Current / $Total) * 100))
    $status = "$Current/$Total"
    if ($Highlight) { $status += " $Highlight" }

    Write-Progress -Id 1 -Activity $Activity -Status $status -PercentComplete $percentComplete
    if ($Current -ge $Total) {
        Write-Progress -Id 1 -Activity $Activity -Completed
    }
}

# Auto-incrementing wrapper around Show-Progress: keeps step numbering sequential
# and lets us report more granular, descriptive sub-steps as they actually happen.
# $Highlight (optional) is appended to the activity text in white, e.g. a filename.
function Step-Progress ($Activity, $Highlight = $null) {
    $script:currentStep++
    Show-Progress $script:currentStep $totalSteps $Activity $Highlight
}

# --- STEP 1: INITIALIZING & PREREQUISITES CHECK ---
$totalSteps = 4
$script:currentStep = 0

# 1. PowerShell Version Check (Requires PS 5.1+, works on both Windows PowerShell and PowerShell 7+)
Step-Progress "Checking PowerShell version"
$psVer = $PSVersionTable.PSVersion
if ($psVer.Major -lt 5 -or ($psVer.Major -eq 5 -and $psVer.Minor -lt 1)) {
    Write-Error "ERROR: This script requires PowerShell 5.1 or higher to run properly."
    Write-Warning "Your current version is: $psVer"
    Write-Warning "Please download and install the latest PowerShell."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# $IsWindows exists only on PowerShell 6+ (Core). On Windows PowerShell 5.1 ("Desktop" edition)
# it is always Windows, since that edition only runs on Windows.
$IsWindowsHost = if ($PSVersionTable.PSEdition -eq 'Desktop') { $true } else { $IsWindows }

# 2. Console Color / ANSI Escape Sequences Support Check
Step-Progress "Enabling console colors (ANSI)"
if ($IsWindowsHost) {
    $MemberSignature = '[DllImport("kernel32.dll", SetLastError=true)] public static extern IntPtr GetStdHandle(int nStdHandle);' +
    '[DllImport("kernel32.dll", SetLastError=true)] public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);' +
    '[DllImport("kernel32.dll", SetLastError=true)] public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);'
    $Type = "Win32.Win32Console" -as [type]
    if (-not $Type) {
        $Type = Add-Type -MemberDefinition $MemberSignature -Name "Win32Console" -Namespace "Win32" -PassThru
    }
    $StdOutHandle = $Type::GetStdHandle(-11) # STD_OUTPUT_HANDLE
    $Mode = 0
    if ($Type::GetConsoleMode($StdOutHandle, [ref]$Mode)) {
        $Mode = $Mode -bor 4 # ENABLE_VIRTUAL_TERMINAL_PROCESSING
        if (-not $Type::SetConsoleMode($StdOutHandle, $Mode)) {
            Write-Warning "WARNING: Failed to enable Virtual Terminal Processing (ANSI colors might not render correctly)."
            Write-Warning "Please use Windows Terminal for the best visual experience: https://aka.ms/terminal"
        }
    }
}

# 3. Braille Character & UTF-8 Console Output Encoding Check
Step-Progress "Checking UTF-8 / Braille rendering support"
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $testBraille = [char]0x28FF
    if ([string]::IsNullOrEmpty($testBraille) -or $testBraille -eq '?') {
        throw "Failed to map Braille Unicode range."
    }
}
catch {
    Write-Error "ERROR: The console does not support UTF-8 or Braille Unicode rendering."
    Write-Warning "To fix this, configure your console font to a compatible TrueType font (e.g., Cascadia Code, JetBrains Mono, or Fira Code)."
    Write-Warning "Get modern developer fonts at: https://www.nerdfonts.com"
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# 4. Internet Connectivity Check (to Yahoo Finance)
Step-Progress "Checking internet connectivity"
$pingTarget = "fc.yahoo.com"
try {
    $ip = [System.Net.Dns]::GetHostAddresses($pingTarget)
    if ($null -eq $ip -or $ip.Count -eq 0) { Write-Warning "Yahoo Finance DNS lookup returned no addresses; the data request will determine whether a refresh is possible." }
}
catch {
    Write-Warning "Yahoo Finance DNS lookup failed: $_. The script will retry after the next data refresh interval."
}

# --------------------------------------------------

while ($true) {
    Set-RefreshCountdownDisplay -Clear

    # Reset progress for each refresh cycle; startup prerequisites run only once.
    $totalSteps = if ($useCsvPortfolio) { 8 } else { 5 }
    $script:currentStep = 0

    # Re-read dimensions in case the console was resized while the script was waiting.
    try {
        if ($Host.UI.RawUI.WindowSize.Width) { $consoleWidth = $Host.UI.RawUI.WindowSize.Width }
    }
    catch {
        Write-Error "ERROR: pfpeek requires an interactive PowerShell host with console window dimensions: $_"
        exit
    }
    if ($consoleWidth -lt $minWidth) {
        $currentSize = Get-ConsoleDimensions
        if ($null -eq $currentSize) { throw "Unable to determine console dimensions while waiting for a resize." }
        Show-ConsoleSizeWarning -RequiredWidth $minWidth -RequiredHeight $currentSize.Height
        continue
    }

    Clear-Host

    # --- STEP 2: BUILD PORTFOLIO (FROM CSV OR MANUAL) ---
    if ($useCsvPortfolio) {
        Step-Progress "Locating Directa CSV export file"

        $searchDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        try {
            $csvPath = Resolve-DirectaCsvPath -ConfiguredPath $configuredCsvPath -SearchDirectory $searchDir
        }
        catch {
            Write-Error "ERROR: $_"
            Write-Warning "Set `$csvPath to the full path of your Directa export, or copy the CSV into the script folder."
            Write-Warning "Alternatively, set `$useCsvPortfolio = `$false to use the manual portfolio."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit
        }
        if (-not (Test-Path $csvPath)) {
            Write-Error "ERROR: CSV file not found: $csvPath"
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit
        }

        Step-Progress "Parsing CSV transactions" (Split-Path $csvPath -Leaf)

        try {
            $movimenti = ConvertFrom-DirectaCsvLines -Lines @(Get-Content -Path $csvPath -Encoding UTF8) -Path $csvPath
        }
        catch {
            Write-Error "ERROR: $_"
            Write-Warning "Ensure the file is an original Directa Trading 'Movimenti' export."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit
        }

        # 3b. Locate (or create) the manual transactions CSV, and merge its rows in.
        # Same column layout/format/delimiter as the Directa export, so it's parsed
        # with the exact same $csvHeaderFields and flows through the same aggregation
        # logic below as the CSV-sourced movimenti.
        $manualCsvPath = if ($null -ne $manualTransactionsCSV -and $manualTransactionsCSV.Trim() -ne "") {
            $manualTransactionsCSV
        }
        else {
            $manualCsvDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
            Join-Path $manualCsvDir "movimentiAggiuntivi.csv"
        }

        if (-not (Test-Path $manualCsvPath)) {
            try {
                # Same header text as the Directa export, so the file can also be
                # opened/edited/appended to with the same conventions as a "real" one.
                $manualCsvHeaderLine = "Data operazione;Data valuta;Tipo operazione;Ticker;Isin;Protocollo;Descrizione;Quantità;Importo euro;Importo Divisa;Divisa;Riferimento ordine"
                $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
                [System.IO.File]::WriteAllText($manualCsvPath, "$manualCsvHeaderLine`r`n", $utf8NoBom)
            }
            catch {
                Write-Error "ERROR: Cannot create manual transactions CSV '$manualCsvPath': $_"
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                exit
            }
        }

        try {
            $manualTransactions = ConvertFrom-DirectaCsvLines `
                -Lines @(Get-Content -Path $manualCsvPath -Encoding UTF8) `
                -Path $manualCsvPath `
                -NormalizeDateSeparators
        }
        catch {
            Write-Error "ERROR: $_"
            Write-Warning "Ensure the file uses the same column layout as the Directa Trading 'Movimenti' export."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit
        }
        if ($manualTransactions.Count -gt 0) { $movimenti = @($movimenti) + @($manualTransactions) }

        Step-Progress "Aggregating portfolio positions"

        # 4. Aggregate transactions per ticker
        $positions = @{}
        foreach ($m in $movimenti) {
            $ticker = $m.Ticker
            if ([string]::IsNullOrWhiteSpace($ticker)) { continue }

            if (-not $positions.ContainsKey($ticker)) {
                $positions[$ticker] = @{ BoughtQty = 0.0; BoughtCost = 0.0; SoldQty = 0.0; SoldValue = 0.0 }
            }

            $qty = ConvertTo-ItalianDouble $m.Quantita
            $importo = ConvertTo-ItalianDouble $m.ImportoEuro

            switch -Wildcard ($m.TipoOperazione) {
                "Acquisto" { $positions[$ticker].BoughtQty += $qty; $positions[$ticker].BoughtCost += [math]::Abs($importo) }
                "Commissioni" { $positions[$ticker].BoughtCost += [math]::Abs($importo) }
                "Vendita" { $positions[$ticker].SoldQty += $qty; $positions[$ticker].SoldValue += [math]::Abs($importo) }
                default { }
            }
        }

        # 5. Build $assets from CSV positions
        $assets = @()
        foreach ($ticker in $positions.Keys) {
            $p = $positions[$ticker]
            $netQty = $p.BoughtQty - $p.SoldQty
            if ($p.BoughtQty -le 0 -or $netQty -le 0.0001) { continue }

            $avgCost = $p.BoughtCost / $p.BoughtQty
            $assets += @{
                Ticker  = "$ticker$tickerSuffix"
                AvgCost = [math]::Round($avgCost, 4)
                Qty     = [math]::Round($netQty, 4)
            }
        }

        # 6. Also consider the assets specified in $manualAssets, adding any ticker not
        # already present from the CSV (a ticker already held is left as reconstructed
        # from the CSV transactions, which is the more accurate/up-to-date source).
        $csvTickersSet = @{}
        foreach ($a in $assets) { $csvTickersSet[$a.Ticker] = $true }
        foreach ($manualAsset in $manualAssets) {
            if (-not $csvTickersSet.ContainsKey($manualAsset.Ticker)) {
                $assets += $manualAsset
                $csvTickersSet[$manualAsset.Ticker] = $true
            }
        }

        if ($assets.Count -eq 0) {
            Write-Error "ERROR: No open positions found in CSV: $csvPath"
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit
        }
    }
    else {
        Step-Progress "Loading manual portfolio configuration"
        $assets = $manualAssets
    }

    # 7. Sort $assets according to the $assetsSorting sequence (not $manualAssets, which is
    # only the static/extra portfolio definition). Tickers not listed in $assetsSorting are
    # placed at the bottom, in their original (CSV aggregation / manual list) order.
    $sortingOrderMap = @{}
    for ($i = 0; $i -lt $assetsSorting.Count; $i++) {
        $sortingOrderMap[$assetsSorting[$i]] = $i
    }
    $assets = @($assets | Sort-Object -Property @{ Expression = {
                if ($sortingOrderMap.ContainsKey($_.Ticker)) { $sortingOrderMap[$_.Ticker] } else { 99999 }
            } 
        })



    # --- PORTFOLIO METRICS (available only when using Directa CSV export) ---
    $csvMetricsAvailable = $useCsvPortfolio
    if ($csvMetricsAvailable) {
        Step-Progress "Computing portfolio metrics"

        $transactionSummary = Get-TransactionSummary -Transactions $movimenti
        $buyCount = $transactionSummary.BuyCount
        $sellCount = $transactionSummary.SellCount
        $totalCommissions = [math]::Abs($transactionSummary.Commissions)
        $totalStampDuty = [math]::Abs($transactionSummary.StampDuty)
        $totalWithholdingTaxes = [math]::Abs($transactionSummary.WithholdingTaxes)
        $totalExpensesSustained = $totalCommissions + $totalStampDuty + $totalWithholdingTaxes

        $totalDeposited = $transactionSummary.Deposited
        $totalWithdrawn = [math]::Abs($transactionSummary.Withdrawn)

        $totalTickersTraded = $positions.Keys.Count
        $totalOpenPositions = $assets.Count
        $totalClosedPositions = $totalTickersTraded - $totalOpenPositions

        # Realized gains: proceeds from sales minus the average-cost basis of the
        # sold shares, summed across every ticker with at least one sale (covers
        # both fully closed positions and partial sells within still-open ones).
        $totalRealizedGain = 0.0
        foreach ($ticker in $positions.Keys) {
            $p = $positions[$ticker]
            if ($p.SoldQty -le 0) { continue }
            $avgCostForSold = if ($p.BoughtQty -gt 0) { $p.BoughtCost / $p.BoughtQty } else { 0.0 }
            $totalRealizedGain += $p.SoldValue - ($p.SoldQty * $avgCostForSold)
        }

        $buyDates = [System.Collections.Generic.List[datetime]]::new()
        $allTradeDates = [System.Collections.Generic.List[datetime]]::new()
        foreach ($m in $movimenti) {
            if ($m.TipoOperazione -eq 'Acquisto' -or $m.TipoOperazione -eq 'Vendita') {
                $d = [DateTime]::MinValue
                if ([DateTime]::TryParseExact($m.DataOperazione, "dd-MM-yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) {
                    $allTradeDates.Add($d)
                    if ($m.TipoOperazione -eq 'Acquisto') { $buyDates.Add($d) }
                }
            }
        }
    
        if ($startDateConfig) {
            $parsedStartDate = [DateTime]::MinValue
            $startDateNormalized = $startDateConfig.Replace('/', '-')
            if ([DateTime]::TryParseExact($startDateNormalized, "dd-MM-yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsedStartDate)) {
                $firstInvestmentDate = $parsedStartDate
            }
            else {
                $firstInvestmentDate = if ($buyDates.Count -gt 0) { ($buyDates | Sort-Object)[0] } else { $null }
            }
        }
        else {
            $firstInvestmentDate = if ($buyDates.Count -gt 0) { ($buyDates | Sort-Object)[0] } else { $null }
        }
        $latestTradeDate = if ($allTradeDates.Count -gt 0) { ($allTradeDates | Sort-Object -Descending)[0] } else { $null }

        # Align $startDateConfig with the oldest transaction found in the CSV when left
        # empty, so the "since" chart/metrics always match the actual investment history.
        if (-not $startDateConfig -and $firstInvestmentDate) {
            $startDateConfig = $firstInvestmentDate.ToString('dd/MM/yyyy')
        }

        # 1. Breakdown investment duration into "y M d", "M d", or "d"
        $investmentDurationFormatted = "0d"
        if ($firstInvestmentDate) {
            $todayDate = [DateTime]::Today
            if ($todayDate -ge $firstInvestmentDate) {
                $years = $todayDate.Year - $firstInvestmentDate.Year
                $months = $todayDate.Month - $firstInvestmentDate.Month
                $days = $todayDate.Day - $firstInvestmentDate.Day

                if ($days -lt 0) {
                    $months--
                    $prevMonth = $todayDate.AddMonths(-1)
                    $days += [DateTime]::DaysInMonth($prevMonth.Year, $prevMonth.Month)
                }
                if ($months -lt 0) {
                    $years--
                    $months += 12
                }

                if ($years -gt 0) {
                    $investmentDurationFormatted = "${years}y ${months}m ${days}d"
                }
                elseif ($firstInvestmentDate.Year -lt $todayDate.Year -or $months -gt 0) {
                    $investmentDurationFormatted = "${months}m ${days}d"
                }
                else {
                    $investmentDurationFormatted = "${days}d"
                }
            }
        }

        # ============================================================
        # CORRECTED CASHFLOW EXTRACTION FOR MWRR
        # ============================================================
        # IMPORTANT: Cash flows are from the INVESTOR's perspective:
        #   - Deposits (Conferimenti): NEGATIVE (money leaves investor)
        #   - Withdrawals (Prelievi): POSITIVE (money returns to investor)
        #   - Costs (Commissioni, Bollo, Rit.etf): NEGATIVE
        #   - Final value: POSITIVE (liquidation proceeds)
        # ============================================================
        $cashflows = New-Object System.Collections.Generic.List[PSCustomObject]

        foreach ($m in $movimenti) {
            $d = [DateTime]::MinValue
            if (-not [DateTime]::TryParseExact($m.DataOperazione, "dd-MM-yyyy", 
                    [System.Globalization.CultureInfo]::InvariantCulture, 
                    [System.Globalization.DateTimeStyles]::None, [ref]$d)) { 
                continue 
            }
        
            $amount = ConvertTo-ItalianDouble $m.ImportoEuro
        
            switch -Wildcard ($m.TipoOperazione) {
                # NEGATIVE: money OUT from investor (deposits into portfolio)
                # Type "External": true external flow between investor and portfolio.
                "Conferimento con bonifico" {
                    $cashflows.Add([PSCustomObject]@{ Date = $d; Amount = - [math]::Abs($amount); Type = "External" })
                }
                # POSITIVE: money IN to investor (withdrawals from portfolio)
                # Type "External": true external flow between investor and portfolio.
                "Prelievo bonifico" {
                    $cashflows.Add([PSCustomObject]@{ Date = $d; Amount = [math]::Abs($amount); Type = "External" })
                }
                # NEGATIVE: costs (fees, stamp duty, taxes)
                # Type "Cost": money leaves the account directly (not a deposit into
                # it), so for TWRR it must NOT be sign-flipped like an external flow.
                "Commissioni" {
                    $cashflows.Add([PSCustomObject]@{ Date = $d; Amount = - [math]::Abs($amount); Type = "Cost" })
                }
                "Bollo portafoglio titoli*" {
                    $cashflows.Add([PSCustomObject]@{ Date = $d; Amount = - [math]::Abs($amount); Type = "Cost" })
                }
                "Rit. etf" {
                    $cashflows.Add([PSCustomObject]@{ Date = $d; Amount = - [math]::Abs($amount); Type = "Cost" })
                }
                # INTERNAL OPERATIONS (no net cash impact - excluded from XIRR)
                "Acquisto" { continue }
                "Vendita" { continue }
                default { continue }
            }
        }

        # Sort cashflows by date
        $cashflows = [System.Collections.Generic.List[PSCustomObject]]($cashflows | Sort-Object Date)

        # Quantity-history for every raw CSV ticker (needed to reconstruct true
        # historical position sizes for TWRR - see Get-TwrrRobust above), and the
        # list of tickers that have since been FULLY closed out (sold to zero) but
        # still need their own historical price series fetched, since they
        # contributed real value to the portfolio for part of the period and are
        # NOT included in $assets (which only holds currently open positions).
        $qtyHistory = Get-QtyHistory $movimenti
        $cashHistory = Get-CashHistory $movimenti
        $openTickersRaw = @($assets | ForEach-Object { $_.Ticker.Substring(0, $_.Ticker.Length - $tickerSuffix.Length) })
        $closedTickers = @($positions.Keys | Where-Object { ($openTickersRaw -notcontains $_) -and ($positions[$_].BoughtQty -gt 0) } | ForEach-Object { "$_$tickerSuffix" })
        # ============================================================
    }



    # 2. ROBUST PARSING OF CONFIG DATE
    $dateParsed = [DateTime]::MinValue
    if (-not [DateTime]::TryParseExact($startDateConfig, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dateParsed)) {
        Write-Error "ERROR: Configured start date ($startDateConfig) is invalid. Use dd/MM/yyyy format."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }
    $customStartUnix = (New-Object DateTimeOffset ($dateParsed)).ToUnixTimeSeconds()
    $nowUnix = (New-Object DateTimeOffset ([DateTime]::UtcNow)).ToUnixTimeSeconds()

    $tickers = (@($assets | ForEach-Object { [uri]::EscapeDataString($_.Ticker) })) -join ","

    # Initialize web session
    $userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    $webSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $webSession.Headers.Add("User-Agent", $userAgent)

    # --- DATA INTERPOLATION FUNCTION (Simplified Linear Interpolation) ---
    function Invoke-InterpolateData ($originalValues, $targetSize) {
        if ($originalValues.Count -lt 2) {
            return , ([double]$originalValues[0]) * $targetSize
        }
    
        $interpolated = [double[]]::new($targetSize)
        $origCount = $originalValues.Count

        for ($i = 0; $i -lt $targetSize; $i++) {
            # NOTE: $i/$targetSize/$origCount are all [int]. When the division below
            # comes out exact (e.g. at i=0 or i=targetSize-1), PowerShell keeps the
            # result as [int] instead of promoting it to [double]. [math]::Floor()/
            # ::Ceiling() are then called with an ambiguous Int32 argument and, per a
            # known PowerShell overload-resolution quirk, silently pick the [decimal]
            # overload instead of the [double] one - poisoning $weight as [decimal].
            # Multiplying that decimal $weight by a [double] delta later then throws
            # "Value was either too large or too small for a Decimal" whenever the
            # delta is huge/NaN/Infinity. Forcing every step to explicit [double]
            # avoids the ambiguous overload entirely, so no decimal ever appears here.
            $percent = [double]$i / [double]($targetSize - 1)
            $exactIdx = $percent * [double]($origCount - 1)

            $lowIdx = [int][math]::Floor([double]$exactIdx)
            $highIdx = [int][math]::Ceiling([double]$exactIdx)
            $weight = [double]$exactIdx - [double]$lowIdx

            $lowVal = [double]$originalValues[$lowIdx]
            $highVal = [double]$originalValues[$highIdx]

            $interpolated[$i] = $lowVal + ($highVal - $lowVal) * $weight
        }
        return $interpolated
    }

    # --- BRAILLE SPARKLINE CHART GENERATOR ---
    function Get-BrailleSparkline ($values, $height = 10, $targetWidth, $visibleRatio = $null, [switch]$Filled) {
        # Beyond plain type-filtering, floating-point NaN/Infinity values (e.g. gaps
        # in historical price data) must also be dropped here: they are still a
        # [valueType] so the old filter let them through, and once inside they
        # propagate through interpolation/min/max/normalization until [math]::Round()
        # tries to cast a NaN to [Int32] and crashes. Filtering them out here means
        # the sparkline simply skips/interpolates over the gap instead of crashing.
        $validValues = @($values | Where-Object {
                if ($_ -isnot [valueType]) { return $false }
                if ($_ -is [double] -or $_ -is [single]) {
                    return -not ([double]::IsNaN($_) -or [double]::IsInfinity($_))
                }
                return $true
            })
        if ($validValues.Count -eq 0) { return "  [No data available to render chart]" }

        $logicalWidth = $targetWidth * 2
    
        $colsToPlot = $logicalWidth
        if ($null -ne $visibleRatio) {
            $colsToPlot = [int][math]::Round($visibleRatio * $logicalWidth)
            if ($colsToPlot -lt 2) { $colsToPlot = 2 }
            if ($colsToPlot -gt $logicalWidth) { $colsToPlot = $logicalWidth }
        }

        $resizedValues = Invoke-InterpolateData -originalValues $validValues -targetSize $colsToPlot

        $min = ($resizedValues | Measure-Object -Minimum).Minimum
        $max = ($resizedValues | Measure-Object -Maximum).Maximum
        $range = $max - $min
        if ($range -eq 0) { $range = 1 }

        $gridHeight = [int]($height * 4)
        $grid = [byte[]]::new($gridHeight * $logicalWidth)

        for ($col = 0; $col -lt $colsToPlot; $col++) {
            $norm = ([double]$resizedValues[$col] - [double]$min) / [double]$range
            # Defensive clamp: guards against any residual NaN/Infinity slipping
            # through (e.g. from an extreme overflow during interpolation) so
            # [math]::Round() never receives a value it can't cast to [Int32].
            if ([double]::IsNaN($norm) -or [double]::IsInfinity($norm)) { $norm = 0.5 }
            $targetY = [int][math]::Round($norm * ($gridHeight - 1))
        
            if ($Filled) {
                for ($row = 0; $row -le $targetY; $row++) {
                    $grid[$row * $logicalWidth + $col] = 1
                }
            }
            else {
                $grid[$targetY * $logicalWidth + $col] = 1
            }
        }

        $brailleRows = [string[]]::new($height)
        for ($r = ($height - 1); $r -ge 0; $r--) {
            $sb = [System.Text.StringBuilder]::new($targetWidth)
            $idx0 = $r * 4
            $idx1 = $r * 4 + 1
            $idx2 = $r * 4 + 2
            $idx3 = $r * 4 + 3

            for ($c = 0; $c -lt $logicalWidth; $c += 2) {
                $c1 = $c
                $c2 = $c + 1

                $p1 = if ($grid[$idx3 * $logicalWidth + $c1] -eq 1) { 1 } else { 0 }
                $p2 = if ($grid[$idx2 * $logicalWidth + $c1] -eq 1) { 2 } else { 0 }
                $p3 = if ($grid[$idx1 * $logicalWidth + $c1] -eq 1) { 4 } else { 0 }
                $p4 = if ($grid[$idx3 * $logicalWidth + $c2] -eq 1) { 8 } else { 0 }
                $p5 = if ($grid[$idx2 * $logicalWidth + $c2] -eq 1) { 16 } else { 0 }
                $p6 = if ($grid[$idx1 * $logicalWidth + $c2] -eq 1) { 32 } else { 0 }
                $p7 = if ($grid[$idx0 * $logicalWidth + $c1] -eq 1) { 64 } else { 0 }
                $p8 = if ($grid[$idx0 * $logicalWidth + $c2] -eq 1) { 128 } else { 0 }

                $codeOffset = $p1 + $p2 + $p3 + $p4 + $p5 + $p6 + $p7 + $p8
                $null = $sb.Append([char](0x2800 + $codeOffset))
            }
            $brailleRows[($height - 1) - $r] = $sb.ToString()
        }
        return $brailleRows
    }

    # --- TIMELINE GENERATORS ---
    function Get-TimelineRow ($timestamps, $targetWidth) {
        if ($null -eq $timestamps -or $timestamps.Count -lt 2) { return " " * $targetWidth }
        $origCount = $timestamps.Count
        $timelineChars = [char[]](" " * $targetWidth)
        $useMonth = $targetWidth -ge 60

        for ($origIdx = 0; $origIdx -lt $origCount; $origIdx++) {
            $epoch = New-Object DateTime 1970, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
            $date = $epoch.AddSeconds($timestamps[$origIdx]).ToLocalTime()

            if ($date.DayOfWeek -eq [DayOfWeek]::Monday) {
                $percent = $origIdx / ($origCount - 1)
                $consoleX = [int][math]::Round($percent * ($targetWidth - 1))

                $dayStr = if ($useMonth) { "$($date.Day)/$($date.Month)" } else { $date.Day.ToString() }
                $len = $dayStr.Length
                $startPos = [math]::Max(0, $consoleX - [int]($len / 2))
            
                if ($startPos + $len -le $targetWidth) {
                    $canWrite = $true
                    for ($k = 0; $k -lt $len; $k++) {
                        if ($timelineChars[$startPos + $k] -ne ' ') { $canWrite = $false; break }
                    }
                    if ($canWrite) {
                        for ($k = 0; $k -lt $len; $k++) { $timelineChars[$startPos + $k] = $dayStr[$k] }
                    }
                }
            }
        }
        return -join $timelineChars
    }

    function Get-IntradayTimelineRow ($targetWidth) {
        $timelineChars = [char[]](" " * $targetWidth)
        $layout = Get-IntradayTimelineLayout -TargetWidth $targetWidth
        foreach ($point in $layout.Labels) {
            $canWrite = $true
            for ($k = 0; $k -lt $point.Label.Length; $k++) {
                if ($timelineChars[$point.Start + $k] -ne ' ') { $canWrite = $false; break }
            }
            if ($canWrite) {
                for ($k = 0; $k -lt $point.Label.Length; $k++) { $timelineChars[$point.Start + $k] = $point.Label[$k] }
            }
        }
        return -join $timelineChars
    }

    function Get-YearlyTimelineRow ($timestamps, $targetWidth) {
        if ($null -eq $timestamps -or $timestamps.Count -lt 2) { return " " * $targetWidth }

        $timelineChars = [char[]](" " * $targetWidth)
        $epoch = New-Object DateTime 1970, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
    
        $monthlyPoints = New-Object System.Collections.Generic.List[PSCustomObject]
        $lastMonthKey = ""

        for ($i = 0; $i -lt $timestamps.Count; $i++) {
            $date = $epoch.AddSeconds($timestamps[$i]).ToLocalTime()
            $monthKey = $date.ToString("MM/yy")
            if ($monthKey -ne $lastMonthKey) {
                $monthlyPoints.Add([PSCustomObject]@{ Index = $i; Label = $monthKey })
                $lastMonthKey = $monthKey
            }
        }

        if ($monthlyPoints.Count -eq 0) { return " " * $targetWidth }

        $requiredSpace = 8
        $maxLabelsAllowed = [math]::Max(1, [int][math]::Floor($targetWidth / $requiredSpace))
        $step = if ($monthlyPoints.Count -gt $maxLabelsAllowed) { [int][math]::Ceiling($monthlyPoints.Count / $maxLabelsAllowed) } else { 1 }

        for ($idx = 0; $idx -lt $monthlyPoints.Count; $idx += $step) {
            $point = $monthlyPoints[$idx]
            $percent = $point.Index / ($timestamps.Count - 1)
            $consoleX = [int][math]::Round($percent * ($targetWidth - 1))
        
            $lbl = $point.Label
            $startPos = [math]::Max(0, $consoleX - [int]($lbl.Length / 2))
            if ($startPos + $lbl.Length -gt $targetWidth) { $startPos = $targetWidth - $lbl.Length }

            $canWrite = $true
            for ($k = 0; $k -lt $lbl.Length; $k++) {
                if ($timelineChars[$startPos + $k] -ne ' ') { $canWrite = $false; break }
            }
            if ($canWrite) {
                for ($k = 0; $k -lt $lbl.Length; $k++) { $timelineChars[$startPos + $k] = $lbl[$k] }
            }
        }
        return -join $timelineChars
    }

    # --- YAHOO CRUMB RETRIEVAL (with GDPR consent fallback, no browser login required) ---
    # Helper: extract the value="" of an <input> tag identified by its name="" attribute,
    # regardless of the order in which the HTML attributes appear.
    function Get-HtmlInputValue {
        param([string]$Html, [string]$Name)
        $tagPattern = "<input\b[^>]*name\s*=\s*[""']$Name[""'][^>]*>"
        $tagMatch = [regex]::Match($Html, $tagPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if (-not $tagMatch.Success) { return $null }
        $valMatch = [regex]::Match($tagMatch.Value, "value\s*=\s*[""']([^""']*)[""']", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($valMatch.Success) { return $valMatch.Groups[1].Value }
        return $null
    }

    # Strategy 1 ("basic"): works out-of-the-box for most US-based IPs.
    function Get-YahooCrumbBasic {
        param($WebSession, [System.Collections.Generic.List[string]]$Diagnostics)
        try {
            $null = Invoke-YahooRequestWithRetry -Request {
                Invoke-WebRequest -Uri "https://fc.yahoo.com" -WebSession $WebSession -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            }
        }
        catch { $Diagnostics.Add("Basic session bootstrap: $(Get-YahooRequestFailureSummary -ErrorRecord $_)") }
        try {
            $resp = Invoke-YahooRequestWithRetry -Request {
                Invoke-WebRequest -Uri "https://query2.finance.yahoo.com/v1/test/getcrumb" -WebSession $WebSession -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            }
            $c = $resp.Content.Trim()
            if ($c -and $c -notmatch "Too Many Requests") { return $c }
            if ($c -match "Too Many Requests") { $Diagnostics.Add("Basic crumb endpoint returned a rate-limit response.") }
            else { $Diagnostics.Add("Basic crumb endpoint returned an empty response.") }
        }
        catch { $Diagnostics.Add("Basic crumb request: $(Get-YahooRequestFailureSummary -ErrorRecord $_)") }
        return $null
    }

    # Strategy 2 ("consent/GDPR"): required for most EU-based IPs (including Italy) the first
    # time a fresh session hits Yahoo, since the "basic" flow gets silently redirected to the
    # cookie-consent page instead of receiving usable cookies. This replicates - purely over
    # HTTP, with no browser and no Yahoo account needed - what happens when a person manually
    # accepts Yahoo's cookie banner.
    function Get-YahooCrumbConsent {
        param($WebSession, [System.Collections.Generic.List[string]]$Diagnostics)
        try {
            $consentResp = Invoke-YahooRequestWithRetry -Request {
                Invoke-WebRequest -Uri "https://guce.yahoo.com/consent" -WebSession $WebSession -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            }
            $html = $consentResp.Content

            $csrfToken = Get-HtmlInputValue -Html $html -Name "csrfToken"
            $sessionId = Get-HtmlInputValue -Html $html -Name "sessionId"
            if (-not $csrfToken -or -not $sessionId) {
                $Diagnostics.Add("Consent page did not contain the expected CSRF/session fields.")
                return $null
            }

            $bodyString = "agree=agree&agree=agree" +
            "&consentUUID=default" +
            "&sessionId=$([uri]::EscapeDataString($sessionId))" +
            "&csrfToken=$([uri]::EscapeDataString($csrfToken))" +
            "&originalDoneUrl=$([uri]::EscapeDataString('https://finance.yahoo.com/'))" +
            "&namespace=yahoo"

            $null = Invoke-WebRequest -Uri "https://consent.yahoo.com/v2/collectConsent?sessionId=$sessionId" `
                -WebSession $WebSession -Method Post -Body $bodyString -ContentType "application/x-www-form-urlencoded" `
                -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            $null = Invoke-YahooRequestWithRetry -Request {
                Invoke-WebRequest -Uri "https://guce.yahoo.com/copyConsent?sessionId=$sessionId" `
                    -WebSession $WebSession -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            }

            $resp = Invoke-YahooRequestWithRetry -Request {
                Invoke-WebRequest -Uri "https://query2.finance.yahoo.com/v1/test/getcrumb" -WebSession $WebSession -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            }
            $c = $resp.Content.Trim()
            if ($c -and $c -notmatch "Too Many Requests") { return $c }
            if ($c -match "Too Many Requests") { $Diagnostics.Add("Consent crumb endpoint returned a rate-limit response.") }
            else { $Diagnostics.Add("Consent crumb endpoint returned an empty response.") }
        }
        catch { $Diagnostics.Add("Consent flow: $(Get-YahooRequestFailureSummary -ErrorRecord $_)") }
        return $null
    }

    function Get-YahooCrumb {
        param($WebSession)
        $diagnostics = [System.Collections.Generic.List[string]]::new()
        $c = Get-YahooCrumbBasic -WebSession $WebSession -Diagnostics $diagnostics
        if (-not $c) { $c = Get-YahooCrumbConsent -WebSession $WebSession -Diagnostics $diagnostics }
        return [pscustomobject]@{ Crumb = $c; Diagnostics = $diagnostics.ToArray() }
    }

    # --- FETCH DATA WITH TIMEOUT AND TRY/CATCH ---
    try {
        Step-Progress "Connecting to Yahoo Finance"

        $crumbResult = Get-YahooCrumb -WebSession $webSession
        $crumb = $crumbResult.Crumb
        if (-not $crumb) {
            $diagnosticSummary = if ($crumbResult.Diagnostics.Count -gt 0) { $crumbResult.Diagnostics -join " | " } else { "No usable response from either strategy." }
            if ($script:lastGoodQuotes.Count -eq 0) {
                throw "Unable to obtain a Yahoo Finance session token (crumb): $diagnosticSummary"
            }
            Write-Warning "Unable to refresh Yahoo Finance quotes ($diagnosticSummary); continuing with cached quotes."
        }
        if ($crumbResult.Diagnostics.Count -gt 0) {
            Write-Verbose "Yahoo crumb retrieval diagnostics: $($crumbResult.Diagnostics -join ' | ')"
        }

        Step-Progress "Downloading real-time quotes"

        $freshQuotesBySymbol = @{}
        if ($crumb) {
            $encodedCrumb = [uri]::EscapeDataString($crumb)
            $apiUrl = "https://query2.finance.yahoo.com/v7/finance/quote?symbols=$tickers&crumb=$encodedCrumb"
            try {
                $response = Invoke-YahooRequestWithRetry -Request {
                    Invoke-RestMethod -Uri $apiUrl -WebSession $webSession -Method Get -TimeoutSec 10
                }
                foreach ($quote in $response.quoteResponse.result) {
                    if ($quote.symbol) { $freshQuotesBySymbol[$quote.symbol] = $quote }
                }
            }
            catch {
                if ($script:lastGoodQuotes.Count -eq 0) { throw }
                $failureSummary = Get-YahooRequestFailureSummary -ErrorRecord $_
                Write-Warning "Live Yahoo quote request failed ($failureSummary); continuing with cached quotes."
            }
        }
        $quoteSnapshot = Resolve-QuoteSnapshot -FreshQuotes $freshQuotesBySymbol -Assets $assets -Cache $script:lastGoodQuotes
        $quotesBySymbol = $quoteSnapshot.Quotes
        $staleQuoteInfo = $quoteSnapshot.Stale
        $missingQuotes = $quoteSnapshot.Missing
        if ($quotesBySymbol.Count -eq 0) {
            throw "Yahoo Finance returned no fresh quotes and no cached quotes are available."
        }
        if ($staleQuoteInfo.Count -gt 0) {
            Write-Warning "Using cached Yahoo quotes for $($staleQuoteInfo.Count) ticker(s); their ages will be shown in the report."
        }
        if ($missingQuotes.Count -gt 0) {
            Write-Warning "Yahoo Finance returned no quote for: $($missingQuotes -join ', '). No cached quote is available; these positions will be omitted from current-value totals."
        }

        Step-Progress "Downloading historical & intraday data"

        # 3. PARALLEL DOWNLOAD OF HISTORICAL DATA VIA RUNSPACE POOL
        # (ForEach-Object -Parallel is a PowerShell 7+ only feature; a runspace pool
        #  works identically on Windows PowerShell 5.1 and PowerShell 7+, out of the box.)
        $today = [DateTime]::Today
        $startTime = (New-Object DateTimeOffset ($today.AddHours(9))).ToUnixTimeSeconds()
        $endTime = (New-Object DateTimeOffset ($today.AddHours(17).AddMinutes(30))).ToUnixTimeSeconds()
        $intradaySteps = 271
        $stepInterval = ($endTime - $startTime) / ($intradaySteps - 1)

        $assetWorkerScript = {
            param($asset, $webSession, $customStartUnix, $nowUnix, $today, $dateParsed, $startTime, $stepInterval, $intradaySteps)

            $ticker = $asset.Ticker
            $avgCost = $asset.AvgCost

            $res = @{
                Ticker           = $ticker
                HistoricalCloses = $null
                GlobalTimestamps = $null
                CustomCloses     = $null
                CustomTimestamps = $null
                AlignedIntraday  = $null
                Warnings         = [System.Collections.Generic.List[string]]::new()
            }

            function Invoke-SafeRest ($uri) {
                for ($attempt = 1; $attempt -le 3; $attempt++) {
                    try {
                        return Invoke-RestMethod -Uri $uri -WebSession $webSession -Method Get -TimeoutSec 10 -ErrorAction Stop
                    }
                    catch {
                        $response = $_.Exception.Response
                        $statusProperty = if ($null -ne $response) { $response.PSObject.Properties["StatusCode"] } else { $null }
                        $statusCode = if ($null -ne $statusProperty) { [int]$statusProperty.Value } else { $null }
                        $isTransient = $null -eq $statusCode -or $statusCode -eq 408 -or $statusCode -eq 429 -or $statusCode -ge 500
                        if (-not $isTransient -or $attempt -eq 3) {
                            $statusLabel = if ($null -ne $statusCode) { " (HTTP $statusCode)" } else { "" }
                            $res.Warnings.Add("Request failed for ${ticker} after $attempt attempt(s)${statusLabel}: $($_.Exception.Message)")
                            return $null
                        }

                        $delaySeconds = [int][math]::Min(8, [math]::Pow(2, $attempt - 1))
                        if ($statusCode -eq 429 -and $null -ne $response) {
                            $headers = $response.PSObject.Properties["Headers"]
                            $retryAfter = if ($null -ne $headers -and $null -ne $headers.Value) { $headers.Value["Retry-After"] } else { $null }
                            $retryAfterSeconds = 0
                            if ([int]::TryParse([string]$retryAfter, [ref]$retryAfterSeconds)) {
                                $delaySeconds = [int][math]::Min(8, [math]::Max(1, $retryAfterSeconds))
                            }
                            else {
                                $retryAfterDate = [datetime]::MinValue
                                if ([datetime]::TryParse([string]$retryAfter, [ref]$retryAfterDate)) {
                                    $delaySeconds = [int][math]::Min(8, [math]::Max(1, [math]::Ceiling(($retryAfterDate - (Get-Date)).TotalSeconds)))
                                }
                            }
                        }
                        Start-Sleep -Seconds $delaySeconds
                    }
                }
            }

            # --- CUSTOM HISTORY ---
            $diffDays = ($today - $dateParsed).TotalDays
            $interval = if ($diffDays -gt 180) { "1wk" } else { "1d" }
            $customChartUrl = "https://query2.finance.yahoo.com/v8/finance/chart/${ticker}?period1=$customStartUnix&period2=$nowUnix&interval=$interval"
            $customChartRes = Invoke-SafeRest $customChartUrl
            if ($customChartRes -and $customChartRes.chart.result) {
                $res.CustomTimestamps = $customChartRes.chart.result[0].timestamp
                $cCloses = $customChartRes.chart.result[0].indicators.quote[0].close
                $cleanedC = [double[]]::new(@($cCloses).Count)
                $lastValidC = $avgCost
                for ($closeIndex = 0; $closeIndex -lt $cleanedC.Count; $closeIndex++) {
                    $p = $cCloses[$closeIndex]
                    if ($null -ne $p -and $p -is [valueType]) { $lastValidC = [double]$p }
                    $cleanedC[$closeIndex] = $lastValidC
                }
                $res.CustomCloses = $cleanedC
            }

            # --- INTRADAY ---
            $intraUrl = "https://query2.finance.yahoo.com/v8/finance/chart/${ticker}?range=1d&interval=2m"
            $intraRes = Invoke-SafeRest $intraUrl
        
            $intraPoints = @{}
            if ($intraRes -and $intraRes.chart.result) {
                $intraTimes = $intraRes.chart.result[0].timestamp
                $intraCloses = $intraRes.chart.result[0].indicators.quote[0].close
                if ($intraTimes) {
                    for ($k = 0; $k -lt $intraTimes.Count; $k++) {
                        $t = $intraTimes[$k]
                        $p = $intraCloses[$k]
                        if ($null -ne $p -and $p -is [valueType]) { $intraPoints[$t] = [double]$p }
                    }
                }
            }
            $sortedIntradayTimes = @($intraPoints.Keys | Sort-Object)
            $intradayIndex = 0

            $alignedIntraday = [double[]]::new($intradaySteps)
            $prevClosePrice = $avgCost
            if ($intraRes -and $intraRes.chart.result) {
                $prevClosePrice = [double]$intraRes.chart.result[0].meta.chartPreviousClose
            }
            $alignedIntraday[0] = $prevClosePrice
            $lastKnownIntraPrice = $prevClosePrice

            for ($step = 1; $step -lt $intradaySteps; $step++) {
                $targetTimestamp = $startTime + (($step - 1) * $stepInterval)
                if ($targetTimestamp -gt $nowUnix) {
                    # Future hour: no real data can exist yet, leave it empty instead of
                    # forward-filling the last known price (which drew a flat, misleading
                    # line from "now" to the end of the day).
                    $alignedIntraday[$step] = [double]::NaN
                    continue
                }
                while ($intradayIndex + 1 -lt $sortedIntradayTimes.Count -and
                    [math]::Abs([double]$sortedIntradayTimes[$intradayIndex + 1] - $targetTimestamp) -lt
                    [math]::Abs([double]$sortedIntradayTimes[$intradayIndex] - $targetTimestamp)) {
                    $intradayIndex++
                }
                if ($sortedIntradayTimes.Count -gt 0) {
                    $bestKey = $sortedIntradayTimes[$intradayIndex]
                    if ([math]::Abs([double]$bestKey - $targetTimestamp) -lt 600) {
                        $lastKnownIntraPrice = $intraPoints[$bestKey]
                    }
                }
                $alignedIntraday[$step] = $lastKnownIntraPrice
            }
            $res.AlignedIntraday = $alignedIntraday

            return $res
        }

        # Run the worker script block once per asset, in parallel, via a runspace pool.
        # Tickers in $closedTickers (fully sold out, so no longer in $assets) ride
        # along on the SAME pool/worker: they still need their own historical price
        # series so Get-TwrrRobust can reconstruct what the portfolio was actually
        # worth while those positions were still open (see $qtyHistory above). They
        # don't need real-time quotes, just the "custom" history the worker already
        # fetches, so a synthetic AvgCost=0 asset is enough.
        $throttleLimit = [math]::Max(1, [math]::Min(8, $assets.Count + $closedTickers.Count))
        $runspacePool = $null
        $pendingJobs = [System.Collections.Generic.List[object]]::new()
        try {
            $runspacePool = [runspacefactory]::CreateRunspacePool(1, $throttleLimit)
            $runspacePool.Open()

            $commonWorkerParameters = @{
                customStartUnix = $customStartUnix
                nowUnix         = $nowUnix
                today           = $today
                dateParsed      = $dateParsed
                startTime       = $startTime
                stepInterval    = $stepInterval
                intradaySteps   = $intradaySteps
            }
            foreach ($asset in $assets) {
                $workerParameters = $commonWorkerParameters.Clone()
                $workerParameters.asset = $asset
                $workerParameters.webSession = New-IndependentWebRequestSession `
                    -SourceSession $webSession -CookieUri ([uri]"https://query2.finance.yahoo.com")
                $pendingJobs.Add((Start-RunspaceScript -RunspacePool $runspacePool -ScriptBlock $assetWorkerScript -Parameters $workerParameters))
            }
            foreach ($closedTicker in $closedTickers) {
                $workerParameters = $commonWorkerParameters.Clone()
                $workerParameters.asset = @{ Ticker = $closedTicker; AvgCost = 0.0 }
                $workerParameters.webSession = New-IndependentWebRequestSession `
                    -SourceSession $webSession -CookieUri ([uri]"https://query2.finance.yahoo.com")
                $pendingJobs.Add((Start-RunspaceScript -RunspacePool $runspacePool -ScriptBlock $assetWorkerScript -Parameters $workerParameters))
            }

            $parallelResults = foreach ($job in $pendingJobs) {
                try {
                    $job.Pipe.EndInvoke($job.Handle)
                }
                catch {
                    Write-Warning "WARNING: A background download task failed: $_"
                }
            }
        }
        finally {
            foreach ($job in $pendingJobs) {
                try {
                    if (-not $job.Handle.IsCompleted) { $job.Pipe.Stop() }
                }
                catch {
                    Write-Warning "Unable to stop a background data task cleanly: $_"
                }
                finally {
                    $job.Pipe.Dispose()
                }
            }
            if ($null -ne $runspacePool) {
                try { $runspacePool.Close() } finally { $runspacePool.Dispose() }
            }
        }

        $customData = @{}
        $rawCustomData = @{}
        $intradayData = @{}
        $globalTimestamps = $null
        $yearlyTimestamps = $null
        $customTimestamps = $null
        # Per-ticker (open AND closed) historical price series, each keeping its
        # own timestamps - used by Get-TwrrRobust to reconstruct true historical
        # portfolio values. Unlike $customData/$customTimestamps below (which only
        # cover currently-open tickers and collapse onto a single shared timeline
        # for the on-screen "since" chart), this is the accurate, per-ticker basis.
        $priceSeriesByTicker = @{}

        foreach ($r in $parallelResults) {
            foreach ($warning in $r.Warnings) { Write-Warning $warning }
            $t = $r.Ticker
            if ($r.AlignedIntraday) { $intradayData[$t] = $r.AlignedIntraday }
            if ($r.CustomCloses -and $r.CustomTimestamps) {
                $priceSeriesByTicker[$t] = @{ Timestamps = $r.CustomTimestamps; Closes = $r.CustomCloses }
                $rawCustomData[$t] = $priceSeriesByTicker[$t]
            }

            if (-not $globalTimestamps -and $r.GlobalTimestamps) { $globalTimestamps = $r.GlobalTimestamps }
            if (-not $yearlyTimestamps -and $r.YearlyTimestamps) { $yearlyTimestamps = $r.YearlyTimestamps }
        }

        $referenceSeries = $null
        foreach ($asset in $assets) {
            if ($rawCustomData.ContainsKey($asset.Ticker) -and
                ($null -eq $referenceSeries -or $rawCustomData[$asset.Ticker].Timestamps.Count -gt $referenceSeries.Timestamps.Count)) {
                $referenceSeries = $rawCustomData[$asset.Ticker]
            }
        }
        if ($null -ne $referenceSeries) {
            $customTimestamps = [long[]]$referenceSeries.Timestamps
            foreach ($asset in $assets) {
                if ($rawCustomData.ContainsKey($asset.Ticker)) {
                    $series = $rawCustomData[$asset.Ticker]
                    $customData[$asset.Ticker] = Get-AlignedHistoricalPrices `
                        -TargetTimestamps $customTimestamps `
                        -SourceTimestamps ([long[]]$series.Timestamps) `
                        -SourcePrices ([double[]]$series.Closes) `
                        -FallbackPrice ([double]$asset.AvgCost)
                }
            }
        }

        $portfolioPrevCloseValue = 0
        foreach ($asset in $assets) {
            $quote = if ($quotesBySymbol.ContainsKey($asset.Ticker)) { $quotesBySymbol[$asset.Ticker] } else { $null }
            if ($quote) {
                $prevPrice = [double]$quote.regularMarketPreviousClose
                $portfolioPrevCloseValue += ($prevPrice * [double]$asset.Qty)
            }
            else {
                $portfolioPrevCloseValue += ([double]$asset.AvgCost * [double]$asset.Qty)
            }
        }

    }
    catch {
        Write-Warning "Yahoo Finance refresh failed: $_"
        Write-Warning "Keeping the script running; another attempt will start in $refreshCountdown minute(s)."
        Wait-ForRefresh -Deadline (Get-Date).AddMinutes($refreshCountdown)
        continue
    }

    Step-Progress "Rendering output"

    $renderScreen = {
        # --- DYNAMIC CONSOLE DIMENSIONS CALCULATION ---
        $consoleHeight = 40
        $renderSize = Get-ConsoleDimensions
        if ($null -ne $renderSize) {
            $consoleWidth = $renderSize.Width
            $consoleHeight = $renderSize.Height
        }
        if ($consoleWidth -lt $minWidth) {
            Show-ConsoleSizeWarning -RequiredWidth $minWidth -RequiredHeight $consoleHeight
            $renderSize = Get-ConsoleDimensions
            if ($null -eq $renderSize) { throw "Unable to determine console dimensions after resizing." }
            $consoleWidth = $renderSize.Width
            $consoleHeight = $renderSize.Height
        }
        $script:lastRenderedConsoleWidth = $consoleWidth
        $script:lastRenderedConsoleHeight = $consoleHeight
        $targetChartWidth = $consoleWidth - 2

        # Minimum height for each chart (in rows) to keep them readable/significant
        $minChartHeight = 3

        # Custom Portfolio Calculations
        $portfolioCustom = [System.Collections.Generic.List[double]]::new()
        $customCount = $customData.Values | ForEach-Object { $_.Count } | Measure-Object -Min
        $minCustom = $customCount.Minimum

        $missingCustomTickers = @($assets | Where-Object { -not $customData.ContainsKey($_.Ticker) } | ForEach-Object { $_.Ticker })
        if ($missingCustomTickers.Count -gt 0) {
            Write-Warning "Custom chart: no historical 'Since $startDateConfig' data returned for: $($missingCustomTickers -join ', '). Their contribution falls back to a flat AvgCost line."
        }

        # TWRR draws on $priceSeriesByTicker, which covers every ticker EVER traded
        # (open positions AND ones since fully closed out) - not just $assets - since
        # closed positions still contributed real value while they were held.
        $allTradedFullTickers = @($positions.Keys | Where-Object { $positions[$_].BoughtQty -gt 0 } | ForEach-Object { "$_$tickerSuffix" })
        $missingTwrrTickers = @($allTradedFullTickers | Where-Object { -not $priceSeriesByTicker.ContainsKey($_) })
        $skipTwrrDueToMissingData = $false
        if ($missingTwrrTickers.Count -gt 0) {
            Write-Warning "TWRR: no historical 'Since $startDateConfig' data returned for: $($missingTwrrTickers -join ', ') (open or since-closed positions). Reconstructed historical portfolio values will miss/understate their contribution; TWRR output will be skipped to avoid misleading results."
            $skipTwrrDueToMissingData = $true
        }

        if ($minCustom -gt 0) {
            for ($idx = 0; $idx -lt $minCustom; $idx++) {
                $idxValue = 0
                foreach ($asset in $assets) {
                    $priceOnIdx = if ($customData.ContainsKey($asset.Ticker)) { $customData[$asset.Ticker][$idx] } else { $asset.AvgCost }
                    $idxValue += ($priceOnIdx * [double]$asset.Qty)
                }
                $portfolioCustom.Add($idxValue)
            }
        }
        else {
            $portfolioCustom.Add(0.0)
        }

        # Intraday Portfolio Calculations
        $portfolioIntraday = [double[]]::new($intradaySteps)
        for ($step = 0; $step -lt $intradaySteps; $step++) {
            $stepValue = 0
            foreach ($asset in $assets) {
                $priceAtStep = if ($intradayData.ContainsKey($asset.Ticker)) { $intradayData[$asset.Ticker][$step] } else { $asset.AvgCost }
                $stepValue += ($priceAtStep * [double]$asset.Qty)
            }
            $portfolioIntraday[$step] = $stepValue
        }

        $localNowUnix = (New-Object DateTimeOffset ([DateTime]::Now)).ToUnixTimeSeconds()
        $currentStepLimit = [int][math]::Floor(($localNowUnix - $startTime) / $stepInterval) + 1
        $currentStepLimit = [math]::Max(1, [math]::Min($currentStepLimit, $intradaySteps - 1))

        $portfolioIntradayFiltered = $portfolioIntraday[0..$currentStepLimit]
        $intradayRatio = $currentStepLimit / ($intradaySteps - 1)

        # Real-time Metrics
        $portfolio = [System.Collections.Generic.List[object]]::new()
        $totalCost = 0
        $totalValue = 0
        $totalDayChange = 0

        foreach ($asset in $assets) {
            $quote = if ($quotesBySymbol.ContainsKey($asset.Ticker)) { $quotesBySymbol[$asset.Ticker] } else { $null }
            if (-not $quote) { continue }

            $currentPrice = [double]$quote.regularMarketPrice
            $prevClose = [double]$quote.regularMarketPreviousClose
            $displayName = $quote.longName

            $cost = [double]$asset.AvgCost * [double]$asset.Qty
            $value = $currentPrice * [double]$asset.Qty
            $gain = $value - $cost
            $gainPct = if ($cost -ne 0) { ($gain / $cost) * 100 } else { 0 }

            $dayChangePerShare = $currentPrice - $prevClose
            $dayChangeTotal = $dayChangePerShare * [double]$asset.Qty
            $dayChangePct = if ($prevClose -ne 0) { ($dayChangePerShare / $prevClose) * 100 } else { 0 }

            $totalCost += $cost
            $totalValue += $value
            $totalDayChange += $dayChangeTotal

            $portfolio.Add([PSCustomObject]@{
                    Ticker         = $asset.Ticker
                    Name           = $displayName
                    CurrentPrice   = $currentPrice
                    AvgCost        = $asset.AvgCost
                    Qty            = $asset.Qty
                    Value          = $value
                    Gain           = $gain
                    GainPct        = $gainPct
                    DayChangeTotal = $dayChangeTotal
                    DayChangePct   = $dayChangePct
                })
        }

        $totalChange = $totalValue - $totalCost
        $totalChangePct = if ($totalCost -ne 0) { ($totalChange / $totalCost) * 100 } else { 0 }
        $totalDayChangePct = if (($totalValue - $totalDayChange) -ne 0) { ($totalDayChange / ($totalValue - $totalDayChange)) * 100 } else { 0 }

        # ============================================================
        # TWRR CALCULATION (computed before $cashflows is mutated below for MWRR)
        # ============================================================
        $twrr = $null
        if ($csvMetricsAvailable -and -not $skipTwrrDueToMissingData) {
            $currentCashValue = Get-CashBalanceAtUnix ((New-Object DateTimeOffset ([DateTime]::Today)).ToUnixTimeSeconds()) $cashHistory
            $twrr = Get-TwrrRobust $cashflows $qtyHistory $priceSeriesByTicker $tickerSuffix $customStartUnix ($totalValue + $currentCashValue) ([DateTime]::Today) $cashHistory
        }
        # ============================================================

        # ============================================================
        # CORRECTED MWRR CALCULATION
        # Add the FINAL portfolio value as a positive cash flow
        # (as if liquidating the entire portfolio at today's value)
        # ============================================================
        $mwrr = $null
        if ($csvMetricsAvailable) {
            $cashflows.Add([PSCustomObject]@{ 
                    Date   = [DateTime]::Today
                    Amount = $totalValue   # POSITIVE: money coming IN to investor
                })
            $mwrr = Get-XirrRobust $cashflows
        }
        # ============================================================

        # --- NUMBER FORMATTING HELPER (Thousand separator "." and fixed decimals) ---
        function Format-NumberLocalized ($value, $decimals = 2) {
            if ($null -eq $value) { return "" }
            $fmt = "F" + $decimals
            $formatted = ([double]$value).ToString($fmt, [System.Globalization.CultureInfo]::InvariantCulture)
            $parts = $formatted.Split('.')
            $intPart = $parts[0]
            if ($parts.Count -gt 1) {
                return "$intPart,$($parts[1])"
            }
            return $intPart
        }

        # --- QUANTITY FORMATTING HELPER ---
        # Integer quantities are printed as plain integers (with thousand separators).
        # Non-integer quantities (e.g. fractional crypto holdings like 0.001197) would
        # be meaningless rounded to 0 decimals, so they're printed as the integer part
        # followed by ",~" to signal "has a non-zero fractional part" without taking up
        # the extra column width a full decimal expansion would need.
        function Format-QtyLocalized ($value) {
            if ($null -eq $value) { return "" }
            $doubleValue = [double]$value
            if ($doubleValue -eq [math]::Truncate($doubleValue)) {
                return Format-NumberLocalized $doubleValue 0
            }
            $intPartStr = ([math]::Truncate($doubleValue)).ToString("N0", [System.Globalization.CultureInfo]::InvariantCulture)
            return "$intPartStr,~"
        }

        # --- CONSOLE PRINTING (BUFFERED OUTPUT TO PREVENT FLICKER) ---
        $ESC = [char]27
        $Green = "$ESC[92m"
        $Red = "$ESC[91m"
        $Gray = "$ESC[90m"
        $White = "$ESC[97m"
        $Orange = "$ESC[33m"
        $Reset = "$ESC[0m"

        function Get-TrendColor ($val) {
            if ($val -gt 0) { return $Green } elseif ($val -lt 0) { return $Red } else { return $Gray }
        }

        function Get-TrendArrow ($val) {
            if ($val -gt 0) { return [char]0x2191 } elseif ($val -lt 0) { return [char]0x2193 } else { return [char]0x2192 }
        }

        # --- ADAPTIVE LAYOUT HELPERS (no truncation, no label/data split) ---

        function Write-WrappedSegments {
            param(
                [System.Text.StringBuilder]$OutputBuffer,
                [string[]]$Segments,
                [int]$ConsoleWidth,
                [string]$Separator = "  ",
                [string]$ContinuationIndent = "  "
            )

            $lines = Get-WrappedSegmentLines -Segments $Segments -ConsoleWidth $ConsoleWidth -Separator $Separator -ContinuationIndent $ContinuationIndent
            foreach ($line in $lines) {
                $null = $OutputBuffer.AppendLine($line)
            }
        }

        # Fixed footprint reserved on the right of every asset block for the sparkline: a constant
        # width (independent of any per-asset text) plus one column of spacing before it. Keeping
        # this constant - rather than deriving it from P/L text length - is what lets the text
        # wrapping and the graph height agree on the same row count.
        $script:AssetGraphWidth = 9
        $script:AssetGraphGap = 1

        function Get-AssetSparklineLayout {
            param([int]$ConsoleWidth)

            # No fixed console-width bands here: the sparkline is always considered enabled, with an
            # unset (dynamic) height. The real height/width - and whether the sparkline ends up
            # disabled after all - is decided purely from measured content (each asset's actual
            # ticker+description length, and the resulting wrapped line count) by the iterative
            # sizing pass and the fallback disable check further down, so it adapts to any console
            # width and any asset name length instead of a hardcoded width->height mapping.
            return [pscustomobject]@{
                Enabled     = $true
                GraphWidth  = $script:AssetGraphWidth
                GraphGap    = $script:AssetGraphGap
                GraphHeight = $null
            }
        }

        # Gets the number of text rows Segments will wrap to once GraphReservedWidth columns are
        # permanently reserved on the right (so the caller can size the sparkline to match).
        function Get-WrappedSegmentRowCount {
            param(
                [string[]]$Segments,
                [int]$ConsoleWidth,
                [string]$Separator = "  ",
                [string]$ContinuationIndent = "  ",
                [int]$GraphReservedWidth = 0
            )

            $textAreaWidth = [math]::Max(1, $ConsoleWidth - $GraphReservedWidth)
            $lines = Get-WrappedSegmentLines -Segments $Segments -ConsoleWidth $textAreaWidth -Separator $Separator -ContinuationIndent $ContinuationIndent
            return [math]::Max(1, $lines.Count)
        }

        # Wraps text Segments exactly like Write-WrappedSegments (a-capo only on whole label+value
        # blocks, never mid-block), but always reserves GraphWidth + GraphGap columns on the right
        # for a sparkline. Because wrapping itself is computed against the narrower text area, the
        # text can never collide with the graph column regardless of console width. The graph's
        # height must equal the resulting row count - call Get-WrappedSegmentRowCount first to size
        # GraphRows accordingly.
        function Write-WrappedSegmentsWithRightGraph {
            param(
                [System.Text.StringBuilder]$OutputBuffer,
                [string[]]$Segments,
                [int]$ConsoleWidth,
                [string]$Separator = "  ",
                [string]$ContinuationIndent = "  ",
                [string[]]$GraphRows,
                [int]$GraphWidth = $script:AssetGraphWidth,
                [int]$GraphGap = $script:AssetGraphGap,
                [int]$GraphColumnStart = 0,
                [string]$GraphColor,
                [string]$Reset,
                [int]$GraphStartLine = 0,
                # Optional shared break points (see Get-SegmentBreaksBefore/Get-LinesFromBreaks). When
                # provided, the same line breaks decided for the whole asset table are replayed here
                # instead of re-measuring against ConsoleWidth, so the printed output matches the row
                # count/graph height already computed for this asset - and lines up with every other
                # asset row in the table instead of wrapping independently.
                [bool[]]$BreaksBefore = $null
            )

            $graphEnabled = $GraphWidth -gt 0
            $textAreaWidth = if ($graphEnabled) { [math]::Max(1, $GraphColumnStart - 1) } else { $ConsoleWidth }
            $lines = if ($null -ne $BreaksBefore) {
                Get-LinesFromBreaks -Segments $Segments -BreaksBefore $BreaksBefore -Separator $Separator -ContinuationIndent $ContinuationIndent
            }
            else {
                Get-WrappedSegmentLines -Segments $Segments -ConsoleWidth $textAreaWidth -Separator $Separator -ContinuationIndent $ContinuationIndent
            }
            if ($lines.Count -eq 0) { $lines = @("") }

            $renderRowCount = if ($graphEnabled) { [math]::Max($lines.Count, $GraphRows.Count + $GraphStartLine) } else { $lines.Count }

            for ($i = 0; $i -lt $renderRowCount; $i++) {
                $textLine = if ($i -lt $lines.Count) { $lines[$i] } else { "" }

                if ($graphEnabled -and $i -ge $GraphStartLine) {
                    $graphRowIndex = $i - $GraphStartLine
                    $graphRow = if ($graphRowIndex -lt $GraphRows.Count) { $GraphRows[$graphRowIndex] } else { [string]([char]0x2800) * $GraphWidth }
                    $textVisLen = Get-VisibleLength $textLine
                    $padToGraphColumn = [math]::Max($GraphGap, $GraphColumnStart - $textVisLen)
                    $coloredGraph = "${GraphColor}${graphRow}${Reset}"
                    $null = $OutputBuffer.AppendLine("$textLine$(' ' * $padToGraphColumn)$coloredGraph")
                }
                else {
                    $null = $OutputBuffer.AppendLine($textLine)
                }
            }
        }

        # --- PORTFOLIO METRICS SUMMARY STRINGS (computed here, before the chart-height sizing
        # below, so Get-WrappedSegmentRowCount can measure their real wrapped row count instead of
        # the sizing logic having to guess it from console-width bands alone) ---
        if ($csvMetricsAvailable) {

            $totalGainAllTime = $totalValue + $totalWithdrawn - $totalDeposited
            $costRatio = Get-CostRatio -TotalGainAllTime $totalGainAllTime -TotalRealizedGain $totalRealizedGain `
                -TotalExpenses $totalExpensesSustained -TotalDeposited $totalDeposited -TotalCost $totalCost
            $costRatioStr = if ($null -ne $costRatio) { "$(Format-NumberLocalized $costRatio 2)%" } else { "N/A" }
            $firstInvestStr = if ($firstInvestmentDate) { $firstInvestmentDate.ToString('dd/MM/yyyy') } else { "N/A" }
            $latestTradeStr = if ($latestTradeDate) { $latestTradeDate.ToString('dd/MM/yyyy') } else { "N/A" }

            $totExpStr = Format-NumberLocalized $totalExpensesSustained 2
            $totComStr = Format-NumberLocalized $totalCommissions 2
            $totStampStr = Format-NumberLocalized $totalStampDuty 2
            $totTaxStr = Format-NumberLocalized $totalWithholdingTaxes 2
            $totDepStr = Format-NumberLocalized $totalDeposited 2
            $totWithStr = Format-NumberLocalized $totalWithdrawn 2

            $m2 = "${Gray}Trades:${Reset} ${White}${buyCount} buy & ${sellCount} sell${Reset}"
            $m3 = "${Gray}Positions:${Reset} ${White}${totalOpenPositions} open & ${totalClosedPositions} closed${Reset}"
            $m4 = "${Gray}First invest:${Reset} ${White}${firstInvestStr}${Reset} (${White}${investmentDurationFormatted} ago${Reset})"
            $m5 = "${Gray}Latest buy/sell:${Reset} ${White}${latestTradeStr}${Reset}"
            $m6 = "${Gray}Deposited:${Reset} ${White}$totDepStr${Reset}  ${Gray}Withdrawn:${Reset} ${White}$totWithStr${Reset}"

            $realizedGainSign = if ($totalRealizedGain -lt 0) { "-" } else { "" }
            $realizedGainStr = $realizedGainSign + (Format-NumberLocalized ([math]::Abs($totalRealizedGain)) 2)
            $realizedGainColor = Get-TrendColor $totalRealizedGain
            $m7 = "Realized: ${realizedGainColor}$(Get-TrendArrow $totalRealizedGain) ${realizedGainStr}${Reset}"

            $m8 = "Expenses: ${Orange}$totExpStr${Reset} (${Orange}${costRatioStr} ${Gray}of realized${Reset})"
            $m9 = "${Gray}CG taxes${Reset} ${Orange}$totTaxStr${Reset} + ${Gray}Commissions${Reset} ${Orange}$totComStr${Reset} + ${Gray}Stamp duty${Reset} ${Orange}$totStampStr${Reset}"
        }



        # F. Individual Assets

        function Get-AssetCategory ($ticker, $name) {
            $t = $ticker.ToUpper()
            $n = $name.ToUpper()

            if ($t -match "GLD" -or $n -match "GOLD" -or $n -match "COMMODITY" -or $n -match "PHYSICAL") {
                return @{ Label = " COMM "; Color = "$ESC[43;30m" }
            }
            if ($t -match "XEON" -or $t -match "LEON" -or $n -match "OVERNIGHT" -or $n -match "LIQUIDITY" -or $n -match "MONEY MARKET") {
                return @{ Label = " MMKT "; Color = "$ESC[100;97m" }
            }
            if ($t -match "EM35" -or $n -match "BOND" -or $n -match "TREASURY" -or $n -match "GOVT" -or $n -match "CORP" -or $n -match "OBBL") {
                return @{ Label = " BOND "; Color = "$ESC[44;97m" }
            }
            if ($t -match "BITCOIN" -or $t -match "BTC" -or $n -match "BITCOIN" -or $n -match "BTC" -or $n -match "CRYPTO") {
                return @{ Label = " CRYP "; Color = "$ESC[46;97m" }
            }
    
            return @{ Label = " EQTY "; Color = "$ESC[41;97m" }
        }

        # Fixed footprint values for the per-asset sparkline: see $script:AssetGraphWidth /
        # $script:AssetGraphGap declared with Get-AssetSparklineLayout above. The actual width/height
        # used is decided further below, only once the table's line count is known (see Pass 1/1b).
        $assetGraphLayoutProbe = Get-AssetSparklineLayout -ConsoleWidth $consoleWidth
        $assetGraphGap = $assetGraphLayoutProbe.GraphGap

        $maxTitleWidth = 0
        foreach ($item in $portfolio) {
            $category = Get-AssetCategory $item.Ticker $item.Name
            $title = "$($category.Color)$($category.Label)$Reset ${White}$($item.Ticker) ${Reset}$($item.Name)${Reset}"
            $maxTitleWidth = [math]::Max($maxTitleWidth, (Get-VisibleLength $title))
        }

        $maxSingleSegmentLen = 0

        # --- Pass 1: build every asset's Segments (text only - no wrapping/graph sizing yet). ---
        $assetSegmentsByItem = [System.Collections.Generic.List[object]]::new()

        foreach ($item in $portfolio) {
            $weight = ($item.Value / $totalValue) * 100
            $gainColor = Get-TrendColor $item.Gain
            $dcColor = Get-TrendColor $item.DayChangeTotal

            $cat = Get-AssetCategory $item.Ticker $item.Name
            $labelFormatted = "$($cat.Color)$($cat.Label)$Reset"

            $qtyStr = Format-QtyLocalized $item.Qty
            $avgCostStr = Format-NumberLocalized $item.AvgCost 2
            $priceStr = Format-NumberLocalized $item.CurrentPrice 2
            $valStr = Format-NumberLocalized $item.Value 2
            $weightStr = Format-NumberLocalized $weight 2

            $gainPctSign = if ($item.GainPct -gt 0) { "+" } else { "" }
            $dayPctSign = if ($item.DayChangePct -gt 0) { "+" } else { "" }
            $gainSign = if ($item.Gain -lt 0) { "-" } else { "" }
            $daySign = if ($item.DayChangeTotal -lt 0) { "-" } else { "" }

            $gainStr = $gainSign + (Format-NumberLocalized ([math]::Abs($item.Gain)) 2)
            $gainPctStr = $gainPctSign + (Format-NumberLocalized $item.GainPct 2)
            $dayStr = $daySign + (Format-NumberLocalized ([math]::Abs($item.DayChangeTotal)) 2)
            $dayPctStr = $dayPctSign + (Format-NumberLocalized $item.DayChangePct 2)

            $graphColor = Get-TrendColor $item.DayChangeTotal

            $titleLine = "$labelFormatted ${White}$($item.Ticker) ${Reset}$($item.Name)${Reset}"
            $rawGainText = "{0} {1,9} ({2,6}%)" -f (Get-TrendArrow $item.Gain), $gainStr, $gainPctStr
            $rawDayText = "{0} {1,8} ({2,6}%)" -f (Get-TrendArrow $item.DayChangeTotal), $dayStr, $dayPctStr
            $gainFormatted = "{0,-21}" -f $rawGainText
            $dayFormatted = "{0,-21}" -f $rawDayText
            $coloredGain = "${gainColor}${gainFormatted}${Reset}"
            $coloredDay = "${dcColor}${dayFormatted}${Reset}"

            $titleWidth = Get-VisibleLength $titleLine
            if ($titleWidth -lt $maxTitleWidth) {
                $titleLine += " " * ($maxTitleWidth - $titleWidth)
            }

            $assetSegments = @(
                $titleLine
                (" ${Gray}Qty:${Reset} {0,4}  ${Gray}Avg:${Reset} {1,8}  ${Gray}Price:${Reset} {2,8}" -f $qtyStr, $avgCostStr, $priceStr)
                (" ${Gray}Value:${Reset} ${White}{0,9} ({1,5}%)${Reset}" -f $valStr, $weightStr)
                " ${Gray}P/L:${Reset} $coloredGain"
                " ${Gray}Day:${Reset} $coloredDay"
            )

            for ($segmentIndex = 0; $segmentIndex -lt $assetSegments.Count; $segmentIndex++) {
                $seg = $assetSegments[$segmentIndex]
                if ($segmentIndex -eq 0) { continue }
                $segLen = Get-VisibleLength $seg
                if ($segLen -gt $maxSingleSegmentLen) { $maxSingleSegmentLen = $segLen }
            }

            $assetSegmentsByItem.Add([pscustomobject]@{
                    Segments   = $assetSegments
                    GraphGap   = $assetGraphGap
                    GraphColor = $graphColor
                    Item       = $item
                })
        }

        $assetSegmentCount = if ($assetSegmentsByItem.Count -gt 0) { $assetSegmentsByItem[0].Segments.Count } else { 0 }

        # Share break points across ticker rows so columns stay aligned while each module remains
        # independently movable to the next line.
        function Get-GlobalBreaksBefore {
            param(
                [object[]]$Entries,
                [int]$TextWidth,
                [int]$SegmentCount
            )
            $breaks = New-Object bool[] ($SegmentCount)
            foreach ($entry in $Entries) {
                $itemBreaksBefore = Get-SegmentBreaksBefore -Segments $entry.Segments -ConsoleWidth $TextWidth -Separator "   " -ContinuationIndent "  "
                for ($breakIdx = 0; $breakIdx -lt $itemBreaksBefore.Count; $breakIdx++) {
                    if ($itemBreaksBefore[$breakIdx]) { $breaks[$breakIdx] = $true }
                }
            }
            return , $breaks
        }

        # Use the graph's minimum width so the text can use as much of the console as possible.
        # Every complete segment (title or labeled value) stays on one line unless the next
        # segment would exceed the available width.
        if ($assetGraphLayoutProbe.Enabled) {
            $globalGraphWidth = $script:AssetGraphWidth
            $textWidth = [math]::Max(1, $consoleWidth - $globalGraphWidth - $assetGraphGap)
            $globalBreaksBefore = Get-GlobalBreaksBefore -Entries $assetSegmentsByItem -TextWidth $textWidth -SegmentCount $assetSegmentCount
            $breakCount = (@($globalBreaksBefore) | Where-Object { $_ }).Count
            $lineCount = 1 + $breakCount
            $graphStartLine = if ($assetSegmentCount -gt 1 -and $globalBreaksBefore[1]) { 1 } else { 0 }
            $globalGraphHeight = [math]::Max(1, $lineCount - $graphStartLine)
        }
        else {
            # No sparkline at all: wrap against the full console width.
            $globalGraphHeight = 0
            $globalBreaksBefore = Get-GlobalBreaksBefore -Entries $assetSegmentsByItem -TextWidth $consoleWidth -SegmentCount $assetSegmentCount
            $graphStartLine = 0
        }

        if (-not $assetGraphLayoutProbe.Enabled) { $globalGraphWidth = 0 }

        $layoutLineCount = 1 + (@($globalBreaksBefore) | Where-Object { $_ }).Count
        if ($layoutLineCount -gt 2) {
            # Omit the combined Qty/Avg/Price block before allowing the ticker layout to use a third line.
            foreach ($entry in $assetSegmentsByItem) {
                $entry.Segments = @($entry.Segments[0], $entry.Segments[2], $entry.Segments[3], $entry.Segments[4])
            }
            $assetSegmentCount = if ($assetSegmentsByItem.Count -gt 0) { $assetSegmentsByItem[0].Segments.Count } else { 0 }

            $maxSingleSegmentLen = 0
            foreach ($entry in $assetSegmentsByItem) {
                for ($segmentIndex = 1; $segmentIndex -lt $entry.Segments.Count; $segmentIndex++) {
                    $segmentLen = Get-VisibleLength $entry.Segments[$segmentIndex]
                    if ($segmentLen -gt $maxSingleSegmentLen) { $maxSingleSegmentLen = $segmentLen }
                }
            }

            $textWidth = if ($assetGraphLayoutProbe.Enabled) {
                [math]::Max(1, $consoleWidth - $globalGraphWidth - $assetGraphGap)
            }
            else {
                $consoleWidth
            }
            $globalBreaksBefore = Get-GlobalBreaksBefore -Entries $assetSegmentsByItem -TextWidth $textWidth -SegmentCount $assetSegmentCount
            $layoutLineCount = 1 + (@($globalBreaksBefore) | Where-Object { $_ }).Count
            if ($assetGraphLayoutProbe.Enabled) {
                $graphStartLine = if ($assetSegmentCount -gt 1 -and $globalBreaksBefore[1]) { 1 } else { 0 }
                $globalGraphHeight = [math]::Max(1, $layoutLineCount - $graphStartLine)
            }
        }

        # --- Pass 2: replay the SAME break points for every asset (Get-LinesFromBreaks), so all rows
        # share an identical line structure, using the sparkline height/width already decided above. ---
        $assetRenderInfos = [System.Collections.Generic.List[object]]::new()

        foreach ($entry in $assetSegmentsByItem) {
            $wrappedTextLines = Get-LinesFromBreaks -Segments $entry.Segments -BreaksBefore $globalBreaksBefore -Separator "   " -ContinuationIndent "  "
            if ($wrappedTextLines.Count -eq 0) { $wrappedTextLines = @("") }

            $assetRenderInfos.Add([pscustomobject]@{
                    Segments         = $entry.Segments
                    GraphWidth       = $globalGraphWidth
                    GraphGap         = $entry.GraphGap
                    GraphHeight      = $globalGraphHeight
                    WrappedTextLines = $wrappedTextLines
                    GraphColor       = $entry.GraphColor
                    Item             = $entry.Item
                })
        }

        $maxTextWidth = 0
        foreach ($renderInfo in $assetRenderInfos) {
            for ($lineIndex = $graphStartLine; $lineIndex -lt $renderInfo.WrappedTextLines.Count; $lineIndex++) {
                $line = $renderInfo.WrappedTextLines[$lineIndex]
                $lineLen = Get-VisibleLength $line
                if ($lineLen -gt $maxTextWidth) { $maxTextWidth = $lineLen }
            }
        }
        $graphColumnStart = if ($maxTextWidth -gt 0) { $maxTextWidth + $script:AssetGraphGap } else { 0 }

        # Clamp the graph column so that text + gap + graph never exceeds $consoleWidth: without
        # this, some combinations of console width and metric values pushed the
        # printed line width past the console's actual width, causing the terminal to auto-wrap
        # mid-line (splitting braille glyphs and injecting stray newlines into the sparkline).
        # Write-WrappedSegmentsWithRightGraph re-wraps each asset's text against the (possibly
        # narrower) clamped column, so this only ever adds text lines - it never truncates or
        # overlaps the graph.
        $maxGraphWidthUsed = 0
        foreach ($renderInfo in $assetRenderInfos) {
            if ($renderInfo.GraphWidth -gt $maxGraphWidthUsed) { $maxGraphWidthUsed = $renderInfo.GraphWidth }
        }
        if ($maxGraphWidthUsed -gt 0) {
            $maxAllowedGraphColumnStart = $consoleWidth - $script:AssetGraphGap - $maxGraphWidthUsed
            if ($graphColumnStart -gt $maxAllowedGraphColumnStart) {
                $graphColumnStart = [math]::Max(1, $maxAllowedGraphColumnStart)
            }
        }

        # Whole segments (individual label/value pairs) are never split, so the graph column
        # can never start before the widest single segment ends - otherwise that segment alone
        # would overflow into (or past) the graph, wrapping the terminal mid-line. If honoring
        # that minimum would leave no usable room for even the narrowest sparkline, drop the
        # per-asset graphs entirely for this render instead of printing a garbled line.
        $minRequiredGraphColumnStart = $maxSingleSegmentLen + 1
        if ($maxGraphWidthUsed -gt 0 -and ($minRequiredGraphColumnStart + $script:AssetGraphGap + $script:AssetGraphWidth) -gt $consoleWidth) {
            foreach ($renderInfo in $assetRenderInfos) { $renderInfo.GraphWidth = 0 }
            $graphColumnStart = 0
        }
        elseif ($graphColumnStart -lt $minRequiredGraphColumnStart) {
            $graphColumnStart = $minRequiredGraphColumnStart
            $maxAllowedGraphWidth = [math]::Max(0, $consoleWidth - $graphColumnStart - $script:AssetGraphGap)
            foreach ($renderInfo in $assetRenderInfos) {
                if ($renderInfo.GraphWidth -gt $maxAllowedGraphWidth) { $renderInfo.GraphWidth = $maxAllowedGraphWidth }
            }
        }

        $quoteStatusSegments = [System.Collections.Generic.List[string]]::new()
        if ($staleQuoteInfo.Count -gt 0 -or $missingQuotes.Count -gt 0) {
            $quoteStatusSegments.Add("${Orange}QUOTE DATA:${Reset}")
            foreach ($ticker in @($staleQuoteInfo.Keys | Sort-Object)) {
                $age = [TimeSpan]::FromSeconds([double]$staleQuoteInfo[$ticker].AgeSeconds)
                $ageLabel = if ($age.TotalDays -ge 1) {
                    "{0:0}d" -f $age.TotalDays
                }
                elseif ($age.TotalHours -ge 1) {
                    "{0:0}h" -f $age.TotalHours
                }
                else {
                    "{0:0}m" -f [math]::Ceiling($age.TotalMinutes)
                }
                $quoteStatusSegments.Add("${Orange}STALE${Reset} ${White}${ticker}${Reset} (${ageLabel})")
            }
            foreach ($ticker in $missingQuotes) {
                $quoteStatusSegments.Add("${Red}MISSING${Reset} ${White}${ticker}${Reset}")
            }
        }

        # --- FIXED ROWS & CHART HEIGHT CALCULATION ---
        # The asset table has already been laid out above using the exact same wrapping/break logic
        # used by the renderer below. Reserve its REAL number of printed lines instead of estimating
        # from console-width bands. Each asset also gets one blank separator line after it.
        $fixedRows = 0
        # Graphs block:
        $fixedRows += 4 # Performance today header + Prev/Now + timeline + separator
        $fixedRows += 3 # Performance since header + timeline + separator
        # Summary block:
        $fixedRows += 2 # Minimum, when console is wide enough and $useCsvPortfolio=$false
        if ($csvMetricsAvailable) {
            # Measured (not guessed) row count: replays the exact same wrapping the E-bis section
            # further below will use to print $m2..$m9, so the reserved space always matches reality
            # - including when it wraps to more than 1 line per group at narrow console widths.
            $summaryGroup1Rows = Get-WrappedSegmentRowCount -Segments @(" $m2", $m3, $m4, $m5, $m6) -ConsoleWidth $consoleWidth -Separator "  " -ContinuationIndent " "
            $summaryGroup2Rows = Get-WrappedSegmentRowCount -Segments @(" $m7", "  $m8", "= $m9") -ConsoleWidth $consoleWidth -Separator " " -ContinuationIndent "                       "
            $fixedRows += $summaryGroup1Rows + $summaryGroup2Rows
        }
        elseif (-not $csvMetricsAvailable) {
            if ($consoleWidth -lt 95 ) { $fixedRows += 1 }
        }
        if ($quoteStatusSegments.Count -gt 0) {
            $fixedRows += Get-WrappedSegmentRowCount -Segments $quoteStatusSegments.ToArray() -ConsoleWidth $consoleWidth -Separator "  " -ContinuationIndent " "
        }
        # Assets table block: use the actual final row count from the shared break points.
        $assetRowsPerItem = if ($assetRenderInfos.Count -gt 0) {
            [math]::Max(1, @($assetRenderInfos[0].WrappedTextLines).Count)
        }
        else {
            0
        }
        $omitAssetSeparators = $false
        if ($assets.Count -gt 0) {
            $fixedRows += $assets.Count * ($assetRowsPerItem + 1)
        }
        # Some space for elaborated PS Prompts at the bottom of the screen:
        $fixedRows += 2

        $availableHeightForCharts = $consoleHeight - $fixedRows
        if ($assets.Count -gt 0 -and $availableHeightForCharts -lt (2 * $minChartHeight)) {
            $fixedRows -= $assets.Count
            $omitAssetSeparators = $true
            $availableHeightForCharts = $consoleHeight - $fixedRows
        }
        $calculatedChartHeight = [math]::Max($minChartHeight, [int][math]::Floor($availableHeightForCharts / 2))

        # Chart Generation with Dynamic Height
        $intradayChartRows = Get-BrailleSparkline -values $portfolioIntradayFiltered -height $calculatedChartHeight -targetWidth $targetChartWidth -visibleRatio $intradayRatio
        $intradayTimelineRow = Get-IntradayTimelineRow -targetWidth $targetChartWidth

        $customChartRows = Get-BrailleSparkline -values $portfolioCustom -height $calculatedChartHeight -targetWidth $targetChartWidth -Filled
        $customTimelineRow = Get-YearlyTimelineRow -timestamps $customTimestamps -targetWidth $targetChartWidth

        $outputBuffer = [System.Text.StringBuilder]::new()

        $null = $outputBuffer.Append("$ESC[2J$ESC[H")

        # A. INTRADAY Chart Block
        $script:renderedTitle = "Portfolio performance $((Get-Date).ToString('dd/MM/yyyy HH:mm'))"
        $null = $outputBuffer.AppendLine("${White}${script:renderedTitle}${Reset}")
        if ($quoteStatusSegments.Count -gt 0) {
            Write-WrappedSegments -OutputBuffer $outputBuffer -Segments $quoteStatusSegments.ToArray() -ConsoleWidth $consoleWidth -Separator "  " -ContinuationIndent " "
        }

        $intraFirstVal = $portfolioPrevCloseValue
        $intraLastVal = if ($portfolioIntradayFiltered.Count -gt 0) { $portfolioIntradayFiltered[-1] } else { 0 }
        $intraChartColor = Get-TrendColor $totalDayChange

        # Prepend "+" to positive percentage values, remove "+" from positive absolute values
        $intraChangePct = $totalDayChangePct
        $pctSign = if ($intraChangePct -gt 0) { "+" } elseif ($intraChangePct -lt 0) { "" } else { "" }
        $pctColor = Get-TrendColor $intraChangePct
        $pctFormatted = "$pctColor$pctSign$(Format-NumberLocalized $intraChangePct 2)%$Reset"
        $cleanPctStr = "$pctSign$(Format-NumberLocalized $intraChangePct 2)%"

        $intraDelta = $totalDayChange
        $deltaColor = Get-TrendColor $intraDelta
        $deltaSign = if ($intraDelta -lt 0) { "-" } else { "+" }
        $deltaValueStr = $deltaSign + (Format-NumberLocalized ([math]::Abs($intraDelta)) 2)

        $cleanNowStr = "Now: " + $deltaValueStr + " (" + $cleanPctStr + ")"

        $labelPrev = "Prev: " + (Format-NumberLocalized $intraFirstVal 2)
        $labelNow = "Now: " + "${deltaColor}${deltaValueStr}${Reset}" + " (" + $pctFormatted + ")"

        $consoleXNow = [int][math]::Round($intradayRatio * ($targetChartWidth - 1))
        $startNow = $consoleXNow - [int]($cleanNowStr.Length / 2)

        $maxStartNow = $targetChartWidth - $cleanNowStr.Length
        if ($startNow -gt $maxStartNow) { $startNow = $maxStartNow }

        $minStartNow = $labelPrev.Length + 2
        if ($startNow -lt $minStartNow) { $startNow = $minStartNow }

        $spacesCount = $startNow - $labelPrev.Length
        if ($spacesCount -lt 1) { $spacesCount = 1 }

        $headerValuesLine = $labelPrev + (" " * $spacesCount) + $labelNow
        $null = $outputBuffer.AppendLine(" ${Reset}${headerValuesLine}${Reset}")

        foreach ($row in $intradayChartRows) {
            $null = $outputBuffer.AppendLine(" ${intraChartColor}${row}${Reset}")
        }
        $null = $outputBuffer.AppendLine(" ${Gray}${intradayTimelineRow}${Reset}")
        $null = $outputBuffer.AppendLine("-" * $consoleWidth)

        # D. Custom Historical Chart Block
        $null = $outputBuffer.AppendLine("${White}Portfolio performance since ${startDateConfig}${Reset}")
        $cFirstVal = if ($portfolioCustom.Count -gt 0) { $portfolioCustom[0] } else { 0 }
        $cLastVal = if ($portfolioCustom.Count -gt 0) { $portfolioCustom[-1] } else { 0 }
        $cChartColor = Get-TrendColor ($cLastVal - $cFirstVal)

        foreach ($row in $customChartRows) {
            $null = $outputBuffer.AppendLine(" ${cChartColor}${row}${Reset}")
        }
        $null = $outputBuffer.AppendLine(" ${Gray}${customTimelineRow}${Reset}")
        $null = $outputBuffer.AppendLine("=" * $consoleWidth)

        # E. Summary Row
        $dcColor = Get-TrendColor $totalDayChange
        $tcColor = Get-TrendColor $totalChange

        if ($csvMetricsAvailable -and $null -ne $twrr) {
            $twrrColor = Get-TrendColor $twrr
            $twrrPctVal = $twrr * 100
            $twrrSign = if ($twrrPctVal -gt 0) { "+" } else { "" }
            $twrrStr = "TWRR: ${twrrColor}${twrrSign}$(Format-NumberLocalized $twrrPctVal 2)% ${Reset}annual"
        }
        else {
            $twrrStr = ""
        }

        if ($csvMetricsAvailable -and $null -ne $mwrr) {
            $mwrrColor = Get-TrendColor $mwrr
            $mwrrPctVal = $mwrr * 100
            $mwrrSign = if ($mwrrPctVal -gt 0) { "+" } else { "" }
            $mwrrStr = "MWRR: ${mwrrColor}${mwrrSign}$(Format-NumberLocalized $mwrrPctVal 2)% ${Reset}annual"
        }
        else {
            $mwrrStr = ""
        }

        $totDayChangeSign = if ($totalDayChange -lt 0) { "-" } else { "" }
        $totDayChangePctSign = if ($totalDayChangePct -gt 0) { "+" } else { "" }
        $totChangeSign = if ($totalChange -lt 0) { "-" } else { "" }
        $totChangePctSign = if ($totalChangePct -gt 0) { "+" } else { "" }

        $totDayChangeStr = $totDayChangeSign + (Format-NumberLocalized ([math]::Abs($totalDayChange)) 2)
        $totDayChangePctStr = $totDayChangePctSign + (Format-NumberLocalized $totalDayChangePct 2)
        $totChangeStr = $totChangeSign + (Format-NumberLocalized ([math]::Abs($totalChange)) 2)
        $totChangePctStr = $totChangePctSign + (Format-NumberLocalized $totalChangePct 2)
        $totValueStr = Format-NumberLocalized $totalValue 2
        $totCostStr = Format-NumberLocalized $totalCost 2

        $s1 = "Day change: ${dcColor}$(Get-TrendArrow $totalDayChange) $totDayChangeStr ($totDayChangePctStr%)$Reset"
        $s2 = $twrrStr
        $s3 = $mwrrStr
        $s4 = "P/L: ${tcColor}$(Get-TrendArrow $totalChange) $totChangeStr ($totChangePctStr%)${Reset} = ${Gray}Value ${White}$totValueStr${Reset} - ${Gray}Cost ${White}$totCostStr${Reset}"

        Write-WrappedSegments -OutputBuffer $outputBuffer -Segments @(" $s1 ", " $s2 ", " $s3 ", " $s4", "$s5") -ConsoleWidth $consoleWidth -Separator " " -ContinuationIndent ""

        # E-bis. Portfolio Metrics (Responsive layout)
        if ($csvMetricsAvailable) {
            Write-WrappedSegments -OutputBuffer $outputBuffer -Segments @(" $m2", $m3, $m4, $m5, $m6) -ConsoleWidth $consoleWidth -Separator "  " -ContinuationIndent " "
            Write-WrappedSegments -OutputBuffer $outputBuffer -Segments @(" $m7", "  $m8", "= $m9") -ConsoleWidth $consoleWidth -Separator " " -ContinuationIndent "                       "
        }


        $null = $outputBuffer.AppendLine("-" * $consoleWidth)

        # --- LOG APPEND: TSV log with header row creation if missing ---
        if (-not $script:isResizeRedraw) {
            try {
                $logPath = Join-Path $targetLogDir "pfpeek-log.tsv"
                $ic = [System.Globalization.CultureInfo]::InvariantCulture

                $logStats = [ordered]@{
                    TotalValue        = $totalValue.ToString('F2', $ic)
                    TotalCost         = $totalCost.ToString('F2', $ic)
                    TotalChange       = $totalChange.ToString('F2', $ic)
                    TotalChangePct    = $totalChangePct.ToString('F2', $ic)
                    TotalDayChange    = $totalDayChange.ToString('F2', $ic)
                    TotalDayChangePct = $totalDayChangePct.ToString('F2', $ic)
                    StaleQuoteTickers = (@($staleQuoteInfo.Keys | Sort-Object) -join ',')
                    MissingQuoteTickers = ($missingQuotes -join ',')
                    TWRRPct           = if ($csvMetricsAvailable -and $null -ne $twrr) { ($twrr * 100).ToString('F2', $ic) } else { "" }
                    MWRRPct           = if ($csvMetricsAvailable -and $null -ne $mwrr) { ($mwrr * 100).ToString('F2', $ic) } else { "" }
                    ExpensesTotal     = if ($csvMetricsAvailable) { $totalExpensesSustained.ToString('F2', $ic) } else { "" }
                    ExpensesOfGainPct = if ($csvMetricsAvailable -and $null -ne $costRatio) { $costRatio.ToString('F2', $ic) } else { "" }
                    Commissions       = if ($csvMetricsAvailable) { $totalCommissions.ToString('F2', $ic) } else { "" }
                    StampDuty         = if ($csvMetricsAvailable) { $totalStampDuty.ToString('F2', $ic) } else { "" }
                    CapitalGainTaxes  = if ($csvMetricsAvailable) { $totalWithholdingTaxes.ToString('F2', $ic) } else { "" }
                    BuyTrades         = if ($csvMetricsAvailable) { $buyCount } else { "" }
                    SellTrades        = if ($csvMetricsAvailable) { $sellCount } else { "" }
                    OpenPositions     = if ($csvMetricsAvailable) { $totalOpenPositions } else { $assets.Count }
                    ClosedPositions   = if ($csvMetricsAvailable) { $totalClosedPositions } else { "" }
                    FirstInvestDate   = if ($csvMetricsAvailable -and $firstInvestmentDate) { $firstInvestmentDate.ToString('dd/MM/yyyy') } else { "" }
                    LatestTradeDate   = if ($csvMetricsAvailable -and $latestTradeDate) { $latestTradeDate.ToString('dd/MM/yyyy') } else { "" }
                    DaysInvested      = if ($csvMetricsAvailable) { $investmentDurationFormatted } else { "" }
                    Deposited         = if ($csvMetricsAvailable) { $totalDeposited.ToString('F2', $ic) } else { "" }
                    Withdrawn         = if ($csvMetricsAvailable) { $totalWithdrawn.ToString('F2', $ic) } else { "" }
                    RealizedGain      = if ($csvMetricsAvailable) { $totalRealizedGain.ToString('F2', $ic) } else { "" }
                }

                $logTimestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                Add-PortfolioLogEntry -Path $logPath -Stats $logStats -Timestamp $logTimestamp
            }
            catch {
                Write-Warning "WARNING: Unable to write log row to ${logPath}: $_"
            }
        }

        foreach ($renderInfo in $assetRenderInfos) {
            $graphRows = @()
            if ($renderInfo.GraphWidth -gt 0) {
                $graphRows = @([string]([char]0x2800) * $renderInfo.GraphWidth) * $renderInfo.GraphHeight
                if ($intradayData.ContainsKey($renderInfo.Item.Ticker)) {
                    # Only consider steps up to "now" (same $currentStepLimit used for the main
                    # portfolio chart), and pass the matching $intradayRatio so the plotted
                    # portion of the graph stops exactly at the current hour, leaving the
                    # remaining (future) hours blank instead of stretching known data across
                    # the full width.
                    $knownSoFar = @($intradayData[$renderInfo.Item.Ticker][0..$currentStepLimit] | Where-Object { -not [double]::IsNaN($_) })
                    if ($knownSoFar.Count -gt 1) {
                        $renderedRows = @(Get-BrailleSparkline -values $knownSoFar -height $renderInfo.GraphHeight -targetWidth $renderInfo.GraphWidth -visibleRatio $intradayRatio)
                        if ($renderedRows.Count -gt 0) { $graphRows = $renderedRows }
                    }
                }
            }

            Write-WrappedSegmentsWithRightGraph -OutputBuffer $outputBuffer -Segments $renderInfo.Segments -ConsoleWidth $consoleWidth -Separator "   " -ContinuationIndent "  " -GraphRows $graphRows -GraphWidth $renderInfo.GraphWidth -GraphGap $renderInfo.GraphGap -GraphColumnStart $graphColumnStart -GraphColor $renderInfo.GraphColor -Reset $Reset -GraphStartLine $graphStartLine -BreaksBefore $globalBreaksBefore

            if (-not $omitAssetSeparators) { $null = $outputBuffer.AppendLine("") }
        }

        $outputText = $outputBuffer.ToString()
        $requiredHeight = [math]::Max(1, [regex]::Matches($outputText, "`n").Count + 1)
        $requiredWidth = [math]::Max($minWidth, $consoleWidth)
        $currentSize = Get-ConsoleDimensions
        if ($null -eq $currentSize) { throw "Unable to determine console dimensions before rendering." }
        if ($currentSize.Width -lt $requiredWidth -or $currentSize.Height -lt $requiredHeight) {
            Show-ConsoleSizeWarning -RequiredWidth $requiredWidth -RequiredHeight $requiredHeight
            $currentSize = Get-ConsoleDimensions
            if ($null -eq $currentSize) { throw "Unable to determine console dimensions after resizing." }
        }
        $script:lastRenderedConsoleWidth = $currentSize.Width
        $script:lastRenderedConsoleHeight = $currentSize.Height
        Write-Host $outputText -NoNewline
    }

    & $renderScreen
    Wait-ForRefresh -Deadline (Get-Date).AddMinutes($refreshCountdown) -OnResize $renderScreen
}

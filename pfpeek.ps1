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
        * Adaptive console layout with dynamic text wrapping.
        * Braille-based sparkline charts for intraday and long-term performance.
        * Comprehensive financial metrics including XIRR/MWRR, expenses, and trades breakdown.
        * Automated TSV logging of portfolio performance statistics on each run.

.NOTES
    AUTHOR:      Raffaele Bianco
    VERSION:     1.78 (2026-08-27)
    BLOG POST:   https://www.raffaelebianco.it/blog/p/pfpeek/
    GITHUB REPO: https://github.com/RaffaeleBianc0/pfpeek
#>



# === CONFIGURATION STARTS HERE =============================================

# Portfolio source:
#   $true  -> rebuilds it by reading the "Movimenti" CSV export from Directa Trading
#   $false -> uses the static $manualAssets list below
$useCsvPortfolio = $true

# --- Option A: Portfolio from CSV (used if $useCsvPortfolio = $true) ---
# - If $csvPath is specified, that exact file is used.
# - If $csvPath is $null, the script automatically searches for the most recent
#   "Movimenti_*.csv" file in the script directory (or current directory).
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

# Custom start date for the overall chart
# (put your first investment date here, or leave empty to use the oldest date read from the CSV):
$startDateConfig = "03/08/2024"

# Logging folder (if $null, the TSV log is saved in the same folder as the script):
$logFolder = $null   # Common alternative: $env:TEMP

# === CONFIGURATION ENDS HERE ===============================================



# --- MINIMUM CONSOLE SIZE CHECK ---
$minWidth = 102
$consoleWidth = $minWidth
if ($Host.UI.RawUI.WindowSize.Width) { $consoleWidth = $Host.UI.RawUI.WindowSize.Width }
if ($consoleWidth -lt $minWidth) {
    Write-Warning "Console window is too narrow ($consoleWidth chars)."
    Write-Warning "Minimum width required: $minWidth chars."
    Write-Warning "Please increase the console width (or decrease the font size) and try again."
    exit
}

# --- LOG FOLDER WRITE ACCESS CHECK ---
$targetLogDir = if ($null -ne $logFolder -and $logFolder.Trim() -ne "") {
    $logFolder
} else {
    if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
}

if (-not (Test-Path $targetLogDir)) {
    try {
        [void][System.IO.Directory]::CreateDirectory($targetLogDir)
    } catch {
        Write-Error "ERROR: Cannot create log directory '$targetLogDir': $_"
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }
}

try {
    $testFile = Join-Path $targetLogDir ([System.IO.Path]::GetRandomFileName())
    [System.IO.File]::WriteAllText($testFile, "test")
    [System.IO.File]::Delete($testFile)
} catch {
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

    $filled = [math]::Min($Total, [math]::Max(0, [math]::Round(($Current / $Total) * $Total)))
    $empty = $Total - $filled

    $filledChar = [char]0x2588
    $emptyChar = [char]0x2591
    $bar = ($filledChar.ToString() * $filled) + ($emptyChar.ToString() * $empty)

    # Pad numbering with a leading 0 only when there's a tens digit to align with (Total > 9)
    $padFormat = if ($Total -gt 9) { "D2" } else { "D1" }

    Write-Host $Current.ToString($padFormat) -ForegroundColor White -NoNewline
    Write-Host "/" -NoNewline
    Write-Host $Total.ToString($padFormat) -ForegroundColor White -NoNewline
    Write-Host " " -NoNewline
    Write-Host "$bar" -ForegroundColor White -NoNewline
    Write-Host " $Activity" -ForegroundColor Cyan -NoNewline
    if ($Highlight) {
        Write-Host " " -NoNewline
        Write-Host $Highlight -ForegroundColor White -NoNewline
    }
    Write-Host ""

    # If completed, print newline
    if ($Current -ge $Total) { Write-Host "" }
}

# Total steps depends on the portfolio source: the CSV branch performs more
# fine-grained sub-steps (locate/parse/aggregate/metrics) than the manual one.
$totalSteps = if ($useCsvPortfolio) { 12 } else { 9 }
$script:currentStep = 0

# Auto-incrementing wrapper around Show-Progress: keeps step numbering sequential
# and lets us report more granular, descriptive sub-steps as they actually happen.
# $Highlight (optional) is appended to the activity text in white, e.g. a filename.
function Step-Progress ($Activity, $Highlight = $null) {
    $script:currentStep++
    Show-Progress $script:currentStep $totalSteps $Activity $Highlight
}

# --- STEP 1: INITIALIZING & PREREQUISITES CHECK ---

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

# Detect whether the script was launched from the Windows Explorer right-click ("RMB")
# context menu (e.g. "Run with PowerShell"), as opposed to an already-open terminal.
# In that case the console window is spawned fresh and closes immediately once the
# script ends, so we reserve 1 extra line and pause with a "Press any key..." prompt
# at the end instead of letting the window vanish.
$script:launchedFromExplorer = $false
if ($IsWindowsHost) {
    try {
        $currentProcessId = $PID
        $parentProcessId = (Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $currentProcessId" -ErrorAction Stop).ParentProcessId
        $parentProcessName = (Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $parentProcessId" -ErrorAction Stop).Name
        if ($parentProcessName -eq "explorer.exe") { $script:launchedFromExplorer = $true }
    } catch {
        $script:launchedFromExplorer = $false
    }
}

# 2. Console Color / ANSI Escape Sequences Support Check
Step-Progress "Enabling console colors (ANSI)"
if ($IsWindowsHost) {
    $MemberSignature = '[DllImport("kernel32.dll", SetLastError=true)] public static extern IntPtr GetStdHandle(int nStdHandle);' +
                       '[DllImport("kernel32.dll", SetLastError=true)] public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);' +
                       '[DllImport("kernel32.dll", SetLastError=true)] public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);'
    $Type = Add-Type -MemberDefinition $MemberSignature -Name "Win32Console" -Namespace "Win32" -PassThru
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
} catch {
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
    if ($null -eq $ip -or $ip.Count -eq 0) { throw "DNS resolution failed." }
} catch {
    Write-Error "ERROR: No internet connection detected or Yahoo Finance is unreachable."
    Write-Warning "Please check your network settings and proxy configuration."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# --------------------------------------------------

# --- STEP 2: BUILD PORTFOLIO (FROM CSV OR MANUAL) ---
if ($useCsvPortfolio) {
    Step-Progress "Locating Directa CSV export file"

    # 1. Locate the CSV file
    if (-not $csvPath) {
        $searchDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        $latestCsv = Get-ChildItem -Path $searchDir -Filter "Movimenti_*.csv" -File -ErrorAction SilentlyContinue |
                     Sort-Object LastWriteTime -Descending |
                     Select-Object -First 1
        if (-not $latestCsv) {
            Write-Error "ERROR: No 'Movimenti_*.csv' file found in '$searchDir' and no `$csvPath specified."
            Write-Warning "Set `$csvPath to the full path of your Directa export, or copy the CSV into the script folder."
            Write-Warning "Alternatively, set `$useCsvPortfolio = `$false to use the manual portfolio."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit
        }
        $csvPath = $latestCsv.FullName
    }
    if (-not (Test-Path $csvPath)) {
        Write-Error "ERROR: CSV file not found: $csvPath"
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }

    # 2. Helper to convert Italian-formatted numbers (e.g., "1.234,56" -> 1234.56)
    function ConvertTo-ItalianDouble ($text) {
        if ([string]::IsNullOrWhiteSpace($text)) { return 0.0 }
        $clean = $text.Trim().Replace(".", "").Replace(",", ".")
        $val = 0.0
        [void][double]::TryParse($clean, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$val)
        return $val
    }

    Step-Progress "Parsing CSV transactions" (Split-Path $csvPath -Leaf)

    # 3. Read CSV and find the transaction header row
    $csvLines = @(Get-Content -Path $csvPath -Encoding UTF8)
    $headerIdx = -1
    for ($i = 0; $i -lt $csvLines.Count; $i++) {
        if ($csvLines[$i].TrimStart([char]0xFEFF) -match '^Data operazione;') { $headerIdx = $i; break }
    }
    if ($headerIdx -lt 0) {
        Write-Error "ERROR: Header 'Data operazione' not found in CSV: $csvPath"
        Write-Warning "Ensure the file is an original Directa Trading 'Movimenti' export."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }

    $csvHeaderFields = @(
        "DataOperazione","DataValuta","TipoOperazione","Ticker","Isin",
        "Protocollo","Descrizione","Quantita","ImportoEuro","ImportoDivisa","Divisa","RiferimentoOrdine"
    )
    $movimenti = $csvLines[($headerIdx + 1)..($csvLines.Count - 1)] |
                 Where-Object { $_.Trim().Length -gt 0 } |
                 ConvertFrom-Csv -Delimiter ";" -Header $csvHeaderFields

    # 3b. Locate (or create) the manual transactions CSV, and merge its rows in.
    # Same column layout/format/delimiter as the Directa export, so it's parsed
    # with the exact same $csvHeaderFields and flows through the same aggregation
    # logic below as the CSV-sourced movimenti.
    $manualCsvPath = if ($null -ne $manualTransactionsCSV -and $manualTransactionsCSV.Trim() -ne "") {
        $manualTransactionsCSV
    } else {
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
        } catch {
            Write-Error "ERROR: Cannot create manual transactions CSV '$manualCsvPath': $_"
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit
        }
    }

    $manualCsvLines = @(Get-Content -Path $manualCsvPath -Encoding UTF8)
    $manualHeaderIdx = -1
    for ($i = 0; $i -lt $manualCsvLines.Count; $i++) {
        if ($manualCsvLines[$i].TrimStart([char]0xFEFF) -match '^Data operazione;') { $manualHeaderIdx = $i; break }
    }
    if ($manualHeaderIdx -lt 0) {
        Write-Error "ERROR: Header 'Data operazione' not found in manual transactions CSV: $manualCsvPath"
        Write-Warning "Ensure the file uses the same column layout as the Directa Trading 'Movimenti' export."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }

    $manualDataLines = @($manualCsvLines[($manualHeaderIdx + 1)..($manualCsvLines.Count - 1)] |
                          Where-Object { $_.Trim().Length -gt 0 })

    if ($manualDataLines.Count -gt 0) {
        $manualTransactions = $manualDataLines | ConvertFrom-Csv -Delimiter ";" -Header $csvHeaderFields

        # NOTE: every date-parsing call site in this script (Get-QtyHistory,
        # Get-CashHistory, the MWRR cashflow extraction, etc.) expects the
        # Directa CSV's native "dd-MM-yyyy" (dash) format, which is also what
        # this file uses since it shares the Directa layout. Normalize "/" to
        # "-" anyway as a defensive fallback (e.g. if the file was hand-edited
        # in Excel and dates got reformatted with slashes): left unparsed, a
        # row would still count toward $positions/$assets (aggregated without
        # date parsing) but stay invisible to $qtyHistory/$cashHistory, making
        # TWRR's historical reconstruction miss/misdate it entirely - producing
        # wildly wrong results.
        $manualMovimenti = foreach ($mt in $manualTransactions) {
            [pscustomobject]@{
                DataOperazione    = ($mt.DataOperazione -replace '/', '-')
                DataValuta        = ($mt.DataValuta -replace '/', '-')
                TipoOperazione    = $mt.TipoOperazione
                Ticker            = $mt.Ticker
                Isin              = $mt.Isin
                Protocollo        = $mt.Protocollo
                Descrizione       = $mt.Descrizione
                Quantita          = $mt.Quantita
                ImportoEuro       = $mt.ImportoEuro
                ImportoDivisa     = $mt.ImportoDivisa
                Divisa            = $mt.Divisa
                RiferimentoOrdine = $mt.RiferimentoOrdine
            }
        }
        $movimenti = @($movimenti) + @($manualMovimenti)
    }

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

        switch ($m.TipoOperazione) {
            "Acquisto"    { $positions[$ticker].BoughtQty  += $qty; $positions[$ticker].BoughtCost += [math]::Abs($importo) }
            "Commissioni" { $positions[$ticker].BoughtCost += [math]::Abs($importo) }
            "Vendita"     { $positions[$ticker].SoldQty    += $qty; $positions[$ticker].SoldValue += [math]::Abs($importo) }
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
} else {
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
} })



# --- PORTFOLIO METRICS (available only when using Directa CSV export) ---
$csvMetricsAvailable = $useCsvPortfolio
if ($csvMetricsAvailable) {
    Step-Progress "Computing portfolio metrics"

    $buyCount = @($movimenti | Where-Object { $_.TipoOperazione -eq 'Acquisto' }).Count
    $sellCount  = @($movimenti | Where-Object { $_.TipoOperazione -eq 'Vendita' }).Count

    function Get-SumAmounts ($transactionType) {
        $tot = ($movimenti | Where-Object { $_.TipoOperazione -eq $transactionType } |
                ForEach-Object { ConvertTo-ItalianDouble $_.ImportoEuro } |
                Measure-Object -Sum).Sum
        if ($null -eq $tot) { return 0.0 }
        return $tot
    }

    $totalCommissions    = [math]::Abs((Get-SumAmounts "Commissioni"))
    $totalStampDuty          = [math]::Abs((Get-SumAmounts "Bollo portafoglio titoli*"))
    $totalWithholdingTaxes       = [math]::Abs((Get-SumAmounts "Rit. etf"))
    $totalExpensesSustained = $totalCommissions + $totalStampDuty + $totalWithholdingTaxes

    $totalDeposited   = Get-SumAmounts "Conferimento con bonifico"
    $totalWithdrawn = [math]::Abs((Get-SumAmounts "Prelievo bonifico"))

    $totalTickersTraded = $positions.Keys.Count
    $totalOpenPositions   = $assets.Count
    $totalClosedPositions   = $totalTickersTraded - $totalOpenPositions

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

    $buyDates = @()
    $allTradeDates = @()
    foreach ($m in $movimenti) {
        if ($m.TipoOperazione -eq 'Acquisto' -or $m.TipoOperazione -eq 'Vendita') {
            $d = [DateTime]::MinValue
            if ([DateTime]::TryParseExact($m.DataOperazione, "dd-MM-yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) {
                $allTradeDates += $d
                if ($m.TipoOperazione -eq 'Acquisto') { $buyDates += $d }
            }
        }
    }
    
    if ($startDateConfig) {
        $parsedStartDate = [DateTime]::MinValue
        $startDateNormalized = $startDateConfig.Replace('/', '-')
        if ([DateTime]::TryParseExact($startDateNormalized, "dd-MM-yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsedStartDate)) {
            $firstInvestmentDate = $parsedStartDate
        } else {
            $firstInvestmentDate = if ($buyDates.Count -gt 0) { ($buyDates | Sort-Object)[0] } else { $null }
        }
    } else {
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
            } elseif ($firstInvestmentDate.Year -lt $todayDate.Year -or $months -gt 0) {
                $investmentDurationFormatted = "${months}m ${days}d"
            } else {
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
        
        switch ($m.TipoOperazione) {
            # NEGATIVE: money OUT from investor (deposits into portfolio)
            # Type "External": true external flow between investor and portfolio.
            "Conferimento con bonifico" {
                $cashflows.Add([PSCustomObject]@{ Date = $d; Amount = -[math]::Abs($amount); Type = "External" })
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
                $cashflows.Add([PSCustomObject]@{ Date = $d; Amount = -[math]::Abs($amount); Type = "Cost" })
            }
            "Bollo portafoglio titoli*" {
                $cashflows.Add([PSCustomObject]@{ Date = $d; Amount = -[math]::Abs($amount); Type = "Cost" })
            }
            "Rit. etf" {
                $cashflows.Add([PSCustomObject]@{ Date = $d; Amount = -[math]::Abs($amount); Type = "Cost" })
            }
            # INTERNAL OPERATIONS (no net cash impact - excluded from XIRR)
            "Acquisto" { continue }
            "Vendita" { continue }
            default { continue }
        }
    }

    # Sort cashflows by date
    $cashflows = [System.Collections.Generic.List[PSCustomObject]]($cashflows | Sort-Object Date)

    # ============================================================
    # ROBUST XIRR CALCULATION FUNCTION
    # ============================================================
    function Get-XirrRobust ($cashflows) {
        if ($cashflows.Count -lt 2) { return $null }
        
        $sorted = @($cashflows | Sort-Object Date)
        $baseDate = $sorted[0].Date

        function Get-Npv ($rate) {
            $sum = 0.0
            foreach ($cf in $sorted) {
                $days = ($cf.Date - $baseDate).TotalDays
                if ($days -eq 0) {
                    $sum += $cf.Amount
                } else {
                    $sum += $cf.Amount / [math]::Pow((1 + $rate), ($days / 365.0))
                }
            }
            return $sum
        }

        # Try different initial guesses
        $guesses = @(-0.99, -0.5, -0.1, 0.0, 0.01, 0.05, 0.1, 0.15, 0.2, 0.3, 0.5, 1.0, 2.0, 5.0)
        
        foreach ($guess in $guesses) {
            $rate = $guess
            for ($iter = 0; $iter -lt 200; $iter++) {
                $npv = Get-Npv $rate
                
                # Derivative approximation
                $epsilon = 0.00001
                $npvEps = Get-Npv ($rate + $epsilon)
                $derivative = ($npvEps - $npv) / $epsilon
                
                if ([math]::Abs($derivative) -lt 1e-12) { break }
                
                $newRate = $rate - $npv / $derivative
                
                # Check for convergence
                if ([math]::Abs($newRate - $rate) -lt 1e-10) {
                    # Validate result
                    if ($newRate -gt -0.999 -and $newRate -lt 100) {
                        return $newRate
                    }
                    break
                }
                
                # Prevent overflow
                if ($newRate -lt -0.999 -or $newRate -gt 100) { break }
                
                $rate = $newRate
            }
        }
        
        # Fallback: try binary search if Newton failed
        $low = -0.99
        $high = 10.0
        $npvLow = Get-Npv $low
        $npvHigh = Get-Npv $high
        
        if ($npvLow * $npvHigh -ge 0) { return $null }
        
        for ($iter = 0; $iter -lt 100; $iter++) {
            $mid = ($low + $high) / 2
            $npvMid = Get-Npv $mid
            
            if ([math]::Abs($npvMid) -lt 1e-8) { return $mid }
            
            if ($npvLow * $npvMid -lt 0) {
                $high = $mid
                $npvHigh = $npvMid
            } else {
                $low = $mid
                $npvLow = $npvMid
            }
        }
        
        return $null
    }
    # ============================================================

    # ============================================================
    # TWRR (TIME-WEIGHTED RATE OF RETURN) CALCULATION
    # Unlike MWRR, TWRR neutralizes the effect of deposit/withdrawal timing by
    # chaining sub-period returns computed between each external cash flow.
    # Portfolio valuations at each cash-flow date are reconstructed from the
    # ACTUAL position size held on that date (built from the Acquisto/Vendita
    # history in the CSV), applied to each ticker's historical close nearest
    # that date - including tickers that have since been fully closed out.
    # This is intentionally more accurate than (and independent of) the
    # "Since <date>" custom chart, which still uses current quantities applied
    # to historical closes as a fast approximation for the on-screen sparkline.
    # ============================================================

    # Builds, per raw CSV ticker (no market suffix), a chronological list of
    # @{ DateUnix; Qty } checkpoints: the cumulative quantity held immediately
    # after each Acquisto/Vendita transaction. Used to look up the exact
    # quantity held on any given historical date (a simple step function).
    function Get-QtyHistory ($transactions) {
        $parsed = New-Object System.Collections.Generic.List[PSCustomObject]
        foreach ($m in $transactions) {
            if ($m.TipoOperazione -ne 'Acquisto' -and $m.TipoOperazione -ne 'Vendita') { continue }
            if ([string]::IsNullOrWhiteSpace($m.Ticker)) { continue }
            $d = [DateTime]::MinValue
            if (-not [DateTime]::TryParseExact($m.DataOperazione, "dd-MM-yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { continue }
            $qty = ConvertTo-ItalianDouble $m.Quantita
            $parsed.Add([PSCustomObject]@{ Ticker = $m.Ticker; Date = $d; Delta = if ($m.TipoOperazione -eq 'Acquisto') { $qty } else { -$qty } })
        }

        $history = @{}
        $running = @{}
        foreach ($t in ($parsed | Sort-Object Date)) {
            if (-not $running.ContainsKey($t.Ticker)) { $running[$t.Ticker] = 0.0 }
            $running[$t.Ticker] += $t.Delta
            if (-not $history.ContainsKey($t.Ticker)) { $history[$t.Ticker] = New-Object System.Collections.Generic.List[PSCustomObject] }
            $history[$t.Ticker].Add([PSCustomObject]@{ DateUnix = (New-Object DateTimeOffset ($t.Date)).ToUnixTimeSeconds(); Qty = $running[$t.Ticker] })
        }
        return $history
    }

    function Get-CashHistory ($transactions) {
        $history = New-Object System.Collections.Generic.List[PSCustomObject]
        $idx = 0
        foreach ($m in $transactions) {
            $d = [DateTime]::MinValue
            if (-not [DateTime]::TryParseExact($m.DataOperazione, "dd-MM-yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) {
                continue
            }

            $amount = ConvertTo-ItalianDouble $m.ImportoEuro
            switch ($m.TipoOperazione) {
                "Conferimento con bonifico" { $delta = [math]::Abs($amount); $type = "External" }
                "Prelievo bonifico"         { $delta = -[math]::Abs($amount); $type = "External" }
                "Commissioni"               { $delta = -[math]::Abs($amount); $type = "Cost" }
                "Bollo portafoglio titoli*" { $delta = -[math]::Abs($amount); $type = "Cost" }
                "Rit. etf"                  { $delta = -[math]::Abs($amount); $type = "Cost" }
                "Acquisto"                  { $delta = -[math]::Abs($amount); $type = "Internal" }
                "Vendita"                   { $delta = [math]::Abs($amount); $type = "Internal" }
                default                      { continue }
            }

            $history.Add([PSCustomObject]@{ DateUnix = (New-Object DateTimeOffset ($d)).ToUnixTimeSeconds(); Delta = $delta; Type = $type; Index = $idx })
            $idx++
        }

        $sorted = $history | Sort-Object DateUnix, Index
        if ($sorted.Count -gt 0) {
            $firstTimestamp = $sorted[0].DateUnix
            $sameDayPoints = $sorted | Where-Object { $_.DateUnix -eq $firstTimestamp }
            $hasExternalOnFirstDate = $sameDayPoints | Where-Object { $_.Type -eq 'External' }
            if (-not $hasExternalOnFirstDate) {
                $netFirstDayDelta = ($sameDayPoints | Measure-Object -Property Delta -Sum).Sum
                if ($netFirstDayDelta -lt 0) {
                    $implicitDeposit = -$netFirstDayDelta
                    $sorted = @([PSCustomObject]@{ DateUnix = $firstTimestamp; Delta = $implicitDeposit; Type = 'External'; Index = -1 }) + $sorted
                }
            }
        }

        return [System.Collections.Generic.List[PSCustomObject]]$sorted
    }

    function Get-CashBalanceAtUnix ($targetUnix, $cashHistory) {
        if (-not $cashHistory -or $cashHistory.Count -eq 0) { return 0.0 }
        $balance = 0.0
        foreach ($point in $cashHistory) {
            if ($point.DateUnix -gt $targetUnix) { break }
            $balance += $point.Delta
        }
        return $balance
    }

    # Quantity actually held in $ticker at $targetUnix: the running total as of
    # the latest checkpoint on or before that date (0 if before the first buy).
    function Get-QtyAtUnix ($tickerHistory, $targetUnix) {
        if (-not $tickerHistory -or $tickerHistory.Count -eq 0) { return 0.0 }
        $qty = 0.0
        foreach ($point in $tickerHistory) {
            if ($point.DateUnix -gt $targetUnix) { break }
            $qty = $point.Qty
        }
        return $qty
    }

    # Finds the index of the closest available timestamp to $targetUnix
    function Get-NearestCustomIndex ($targetUnix, $timestamps) {
        if (-not $timestamps -or $timestamps.Count -eq 0) { return -1 }
        $bestIdx = 0
        $bestDiff = [double]::MaxValue
        for ($i = 0; $i -lt $timestamps.Count; $i++) {
            $diff = [math]::Abs([double]$timestamps[$i] - $targetUnix)
            if ($diff -lt $bestDiff) { $bestDiff = $diff; $bestIdx = $i }
        }
        return $bestIdx
    }

    # True historical portfolio value at $targetUnix: sums, over every ticker
    # ever traded (open or since closed), the quantity actually held on that
    # date times its closing price nearest that date.
    function Get-HistoricalPortfolioValue ($targetUnix, $qtyHistory, $priceSeriesByTicker, $tickerSuffix, $cashHistory) {
        $total = 0.0
        foreach ($fullTicker in $priceSeriesByTicker.Keys) {
            $rawTicker = if ($fullTicker.EndsWith($tickerSuffix)) { $fullTicker.Substring(0, $fullTicker.Length - $tickerSuffix.Length) } else { $fullTicker }
            $qty = Get-QtyAtUnix $qtyHistory[$rawTicker] $targetUnix
            if ($qty -le 0) { continue }
            $series = $priceSeriesByTicker[$fullTicker]
            $idx = Get-NearestCustomIndex $targetUnix $series.Timestamps
            if ($idx -lt 0 -or $idx -ge $series.Closes.Count) { continue }
            $total += ([double]$series.Closes[$idx]) * $qty
        }
        $total += Get-CashBalanceAtUnix $targetUnix $cashHistory
        return $total
    }

    function Get-TwrrRobust ($externalFlows, $qtyHistory, $priceSeriesByTicker, $tickerSuffix, $windowStartUnix, $endValue, $endDate, $cashHistory) {
        if (-not $priceSeriesByTicker -or $priceSeriesByTicker.Keys.Count -eq 0) {
            Write-Warning "TWRR: skipped - no historical price series available for any traded ticker. Likely a failed/rate-limited Yahoo Finance download."
            return $null
        }

        # If the configured start date predates the actual first cash movement
        # (deposit/buy), there is nothing invested yet at $windowStartUnix and
        # the reconstructed value is legitimately zero. Clamp the window start
        # to the real first activity date instead of failing outright.
        if ($cashHistory -and $cashHistory.Count -gt 0 -and $cashHistory[0].DateUnix -gt $windowStartUnix) {
            $windowStartUnix = $cashHistory[0].DateUnix
        }

        $periodStartValue = Get-HistoricalPortfolioValue $windowStartUnix $qtyHistory $priceSeriesByTicker $tickerSuffix $cashHistory
        if ($periodStartValue -le 0) {
            Write-Warning "TWRR: skipped - reconstructed initial portfolio value is zero or negative ($periodStartValue)."
            return $null
        }
        $periodStartDate = ([DateTimeOffset]::FromUnixTimeSeconds([int64]$windowStartUnix)).LocalDateTime.Date

        $cumulativeFactor = 1.0
        $sortedFlows = @($externalFlows | Sort-Object Date)

        foreach ($cf in $sortedFlows) {
            $cfUnix = (New-Object DateTimeOffset ($cf.Date)).ToUnixTimeSeconds()
            if ($cfUnix -lt $windowStartUnix) { continue } # flow predates the historical window; can't isolate it

            $valueAtFlow = Get-HistoricalPortfolioValue $cfUnix $qtyHistory $priceSeriesByTicker $tickerSuffix $cashHistory
            $portfolioFlow = if ($cf.Type -eq "Cost") { $cf.Amount } else { -$cf.Amount }
            $valueBeforeFlow = $valueAtFlow - $portfolioFlow
            $ratio = if ($periodStartValue -ne 0) { $valueBeforeFlow / $periodStartValue } else { [double]::NaN }
            $cumulativeFactor *= $ratio

            # External flows (deposits/withdrawals) move money between investor and
            # portfolio, so the portfolio-side impact is the OPPOSITE sign of the
            # investor-side amount (a deposit is negative for the investor but a
            # positive inflow for the portfolio).
            # Costs (commissions, stamp duty, withholding tax) are NOT external to
            # the portfolio: they leave the account directly, reducing its value in
            # the SAME direction as the investor-side amount. Sign-flipping them
            # here (as an earlier version did) makes every cost event look like a
            # deposit, inflating the value base and dragging TWRR down artificially.
            $periodStartValue = $valueBeforeFlow + $portfolioFlow
            if ($periodStartValue -le 0) {
                Write-Warning "TWRR: skipped - sub-period value base hit zero/negative ($periodStartValue) after cash flow on $($cf.Date.ToString('dd/MM/yyyy')) (Type: $($cf.Type), Amount: $($cf.Amount)). Can't chain returns past a wipeout."
                return $null
            }
        }

        if ($periodStartValue -le 0) {
            Write-Warning "TWRR: skipped - final sub-period value base is zero or negative ($periodStartValue)."
            return $null
        }
        $cumulativeFactor *= ($endValue / $periodStartValue)
        if ($cumulativeFactor -le 0) {
            Write-Warning "TWRR: skipped - cumulative return factor is zero or negative ($cumulativeFactor); current total value / last value base gave an invalid ratio."
            return $null
        }

        $totalDays = ($endDate - $periodStartDate).TotalDays
        if ($totalDays -le 0) { return ($cumulativeFactor - 1.0) }

        return ([math]::Pow($cumulativeFactor, 365.0 / $totalDays) - 1.0)
    }

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

$tickers = ($assets.Ticker) -join ","

# Initialize web session
$userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
$webSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$webSession.Headers.Add("User-Agent", $userAgent)

# --- DATA INTERPOLATION FUNCTION (Simplified Linear Interpolation) ---
function Invoke-InterpolateData ($originalValues, $targetSize) {
    if ($originalValues.Count -lt 2) {
        return ,([double]$originalValues[0]) * $targetSize
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
        } else {
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
    $points = @(
        @{ Label = "Prev"; Hour = 9; Min = 0; IsPrev = $true }
        @{ Label = "09";   Hour = 9; Min = 0 }
        @{ Label = "10";   Hour = 10; Min = 0 }
        @{ Label = "11";   Hour = 11; Min = 0 }
        @{ Label = "12";   Hour = 12; Min = 0 }
        @{ Label = "13";   Hour = 13; Min = 0 }
        @{ Label = "14";   Hour = 14; Min = 0 }
        @{ Label = "15";   Hour = 15; Min = 0 }
        @{ Label = "16";   Hour = 16; Min = 0 }
        @{ Label = "17";   Hour = 17; Min = 0 }
        @{ Label = "17:30"; Hour = 17; Min = 30 }
    )
    
    foreach ($p in $points) {
        if ($p.IsPrev) {
            $percent = 0.0
        } else {
            $totalMinutesStart = 9 * 60
            $totalMinutesEnd = (17 * 60) + 30
            $currentMinutes = $p.Hour * 60 + $p.Min
            $percent = 0.12 + (0.88 * (($currentMinutes - $totalMinutesStart) / ($totalMinutesEnd - $totalMinutesStart)))
        }
        
        $lblText = $p.Label
        $consoleX = [int][math]::Round($percent * ($targetWidth - 1))
        
        $startPos = [math]::Max(0, $consoleX - [int]($lblText.Length / 2))
        if ($startPos + $lblText.Length -gt $targetWidth) { $startPos = $targetWidth - $lblText.Length }
        
        $canWrite = $true
        for ($k = 0; $k -lt $lblText.Length; $k++) {
            if ($timelineChars[$startPos + $k] -ne ' ') { $canWrite = $false; break }
        }
        if ($canWrite) {
            for ($k = 0; $k -lt $lblText.Length; $k++) { $timelineChars[$startPos + $k] = $lblText[$k] }
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
    param($WebSession)
    try {
        $null = Invoke-WebRequest -Uri "https://fc.yahoo.com" -WebSession $WebSession -UseBasicParsing -TimeoutSec 10 -ErrorAction SilentlyContinue
    } catch {}
    try {
        $resp = Invoke-WebRequest -Uri "https://query2.finance.yahoo.com/v1/test/getcrumb" -WebSession $WebSession -UseBasicParsing -TimeoutSec 10
        $c = $resp.Content.Trim()
        if ($c -and $c -notmatch "Too Many Requests") { return $c }
    } catch {}
    return $null
}

# Strategy 2 ("consent/GDPR"): required for most EU-based IPs (including Italy) the first
# time a fresh session hits Yahoo, since the "basic" flow gets silently redirected to the
# cookie-consent page instead of receiving usable cookies. This replicates - purely over
# HTTP, with no browser and no Yahoo account needed - what happens when a person manually
# accepts Yahoo's cookie banner.
function Get-YahooCrumbConsent {
    param($WebSession)
    try {
        $consentResp = Invoke-WebRequest -Uri "https://guce.yahoo.com/consent" -WebSession $WebSession -UseBasicParsing -TimeoutSec 10
        $html = $consentResp.Content

        $csrfToken = Get-HtmlInputValue -Html $html -Name "csrfToken"
        $sessionId = Get-HtmlInputValue -Html $html -Name "sessionId"
        if (-not $csrfToken -or -not $sessionId) { return $null }

        $bodyString = "agree=agree&agree=agree" +
            "&consentUUID=default" +
            "&sessionId=$([uri]::EscapeDataString($sessionId))" +
            "&csrfToken=$([uri]::EscapeDataString($csrfToken))" +
            "&originalDoneUrl=$([uri]::EscapeDataString('https://finance.yahoo.com/'))" +
            "&namespace=yahoo"

        $null = Invoke-WebRequest -Uri "https://consent.yahoo.com/v2/collectConsent?sessionId=$sessionId" `
            -WebSession $WebSession -Method Post -Body $bodyString -ContentType "application/x-www-form-urlencoded" `
            -UseBasicParsing -TimeoutSec 10
        $null = Invoke-WebRequest -Uri "https://guce.yahoo.com/copyConsent?sessionId=$sessionId" `
            -WebSession $WebSession -UseBasicParsing -TimeoutSec 10

        $resp = Invoke-WebRequest -Uri "https://query2.finance.yahoo.com/v1/test/getcrumb" -WebSession $WebSession -UseBasicParsing -TimeoutSec 10
        $c = $resp.Content.Trim()
        if ($c -and $c -notmatch "Too Many Requests") { return $c }
    } catch {}
    return $null
}

function Get-YahooCrumb {
    param($WebSession)
    $c = Get-YahooCrumbBasic -WebSession $WebSession
    if (-not $c) { $c = Get-YahooCrumbConsent -WebSession $WebSession }
    return $c
}

# --- FETCH DATA WITH TIMEOUT AND TRY/CATCH ---
try {
    Step-Progress "Connecting to Yahoo Finance"

    $crumb = Get-YahooCrumb -WebSession $webSession
    if (-not $crumb) {
        Write-Error "ERROR: Unable to obtain a Yahoo Finance session token (crumb), even after attempting the GDPR consent fallback."
        Write-Warning "Yahoo Finance may be temporarily blocking automated requests from this IP. Try again in a few minutes."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }

    Step-Progress "Downloading real-time quotes"

    $apiUrl = "https://query2.finance.yahoo.com/v7/finance/quote?symbols=$tickers&crumb=$crumb"
    $response = Invoke-RestMethod -Uri $apiUrl -WebSession $webSession -Method Get -TimeoutSec 10
    $quotes = $response.quoteResponse.result

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
            Ticker = $ticker
            HistoricalCloses = $null
            GlobalTimestamps = $null
            CustomCloses = $null
            CustomTimestamps = $null
            AlignedIntraday = $null
        }

        function Invoke-SafeRest ($uri) {
            try { return Invoke-RestMethod -Uri $uri -WebSession $webSession -Method Get -TimeoutSec 10 } catch { return $null }
        }

        # --- CUSTOM HISTORY ---
        $diffDays = ($today - $dateParsed).TotalDays
        $interval = if ($diffDays -gt 180) { "1wk" } else { "1d" }
        $customChartUrl = "https://query2.finance.yahoo.com/v8/finance/chart/${ticker}?period1=$customStartUnix&period2=$nowUnix&interval=$interval"
        $customChartRes = Invoke-SafeRest $customChartUrl
        if ($customChartRes -and $customChartRes.chart.result) {
            $res.CustomTimestamps = $customChartRes.chart.result[0].timestamp
            $cCloses = $customChartRes.chart.result[0].indicators.quote[0].close
            $cleanedC = @()
            $lastValidC = $avgCost
            foreach ($p in $cCloses) {
                if ($null -ne $p -and $p -is [valueType]) { $lastValidC = [double]$p }
                $cleanedC += $lastValidC
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
            $bestKey = $null
            $minDiff = 600
            foreach ($key in $intraPoints.Keys) {
                $diff = [math]::Abs($key - $targetTimestamp)
                if ($diff -lt $minDiff) {
                    $minDiff = $diff
                    $bestKey = $key
                }
            }
            if ($null -ne $bestKey) { $lastKnownIntraPrice = $intraPoints[$bestKey] }
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
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $throttleLimit)
    $runspacePool.Open()

    $pendingJobs = @()
    foreach ($asset in $assets) {
        $psInstance = [powershell]::Create()
        $psInstance.RunspacePool = $runspacePool
        [void]$psInstance.AddScript($assetWorkerScript).AddParameters(@{
            asset            = $asset
            webSession       = $webSession
            customStartUnix  = $customStartUnix
            nowUnix          = $nowUnix
            today            = $today
            dateParsed       = $dateParsed
            startTime        = $startTime
            stepInterval     = $stepInterval
            intradaySteps    = $intradaySteps
        })
        $pendingJobs += [PSCustomObject]@{
            Pipe   = $psInstance
            Handle = $psInstance.BeginInvoke()
        }
    }
    foreach ($closedTicker in $closedTickers) {
        $psInstance = [powershell]::Create()
        $psInstance.RunspacePool = $runspacePool
        [void]$psInstance.AddScript($assetWorkerScript).AddParameters(@{
            asset            = @{ Ticker = $closedTicker; AvgCost = 0.0 }
            webSession       = $webSession
            customStartUnix  = $customStartUnix
            nowUnix          = $nowUnix
            today            = $today
            dateParsed       = $dateParsed
            startTime        = $startTime
            stepInterval     = $stepInterval
            intradaySteps    = $intradaySteps
        })
        $pendingJobs += [PSCustomObject]@{
            Pipe   = $psInstance
            Handle = $psInstance.BeginInvoke()
        }
    }

    $parallelResults = foreach ($job in $pendingJobs) {
        try {
            $job.Pipe.EndInvoke($job.Handle)
        } catch {
            Write-Warning "WARNING: A background download task failed: $_"
        } finally {
            $job.Pipe.Dispose()
        }
    }

    $runspacePool.Close()
    $runspacePool.Dispose()

    $customData     = @{}
    $intradayData   = @{}
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
        $t = $r.Ticker
        if ($r.CustomCloses)     { $customData[$t]     = $r.CustomCloses }
        if ($r.AlignedIntraday)  { $intradayData[$t]   = $r.AlignedIntraday }
        if ($r.CustomCloses -and $r.CustomTimestamps) {
            $priceSeriesByTicker[$t] = @{ Timestamps = $r.CustomTimestamps; Closes = $r.CustomCloses }
        }

        if (-not $globalTimestamps -and $r.GlobalTimestamps) { $globalTimestamps = $r.GlobalTimestamps }
        if (-not $yearlyTimestamps -and $r.YearlyTimestamps) { $yearlyTimestamps = $r.YearlyTimestamps }
        if (-not $customTimestamps -and $r.CustomTimestamps) { $customTimestamps = $r.CustomTimestamps }
    }

    $portfolioPrevCloseValue = 0
    foreach ($asset in $assets) {
        $quote = $quotes | Where-Object { $_.symbol -eq $asset.Ticker }
        if ($quote) {
            $prevPrice = [double]$quote.regularMarketPreviousClose
            $portfolioPrevCloseValue += ($prevPrice * [double]$asset.Qty)
        } else {
            $portfolioPrevCloseValue += ([double]$asset.AvgCost * [double]$asset.Qty)
        }
    }

} catch {
    Write-Error "Error retrieving data from Yahoo Finance: $_"
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

Step-Progress "Rendering output"

# --- DYNAMIC CONSOLE DIMENSIONS CALCULATION ---
$consoleHeight = 40
if ($Host.UI.RawUI.WindowSize.Height) { $consoleHeight = $Host.UI.RawUI.WindowSize.Height }
$targetChartWidth = $consoleWidth - 2

# Minimum height for each chart (in rows) to keep them readable/significant
$minChartHeight = 3

# Custom Portfolio Calculations
$portfolioCustom = @()
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
        $portfolioCustom += $idxValue
    }
} else {
    $portfolioCustom = @(0)
}

# Intraday Portfolio Calculations
$portfolioIntraday = @()
for ($step = 0; $step -lt $intradaySteps; $step++) {
    $stepValue = 0
    foreach ($asset in $assets) {
        $priceAtStep = if ($intradayData.ContainsKey($asset.Ticker)) { $intradayData[$asset.Ticker][$step] } else { $asset.AvgCost }
        $stepValue += ($priceAtStep * [double]$asset.Qty)
    }
    $portfolioIntraday += $stepValue
}

$localNowUnix = (New-Object DateTimeOffset ([DateTime]::Now)).ToUnixTimeSeconds()
$currentStepLimit = [int][math]::Floor(($localNowUnix - $startTime) / $stepInterval) + 1
$currentStepLimit = [math]::Max(1, [math]::Min($currentStepLimit, $intradaySteps - 1))

$portfolioIntradayFiltered = $portfolioIntraday[0..$currentStepLimit]
$intradayRatio = $currentStepLimit / ($intradaySteps - 1)

# Real-time Metrics
$portfolio = @()
$totalCost = 0
$totalValue = 0
$totalDayChange = 0

foreach ($asset in $assets) {
    $quote = $quotes | Where-Object { $_.symbol -eq $asset.Ticker }
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

    $portfolio += [PSCustomObject]@{
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
    }
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
        Date = [DateTime]::Today
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

# Visible length of a string, ignoring any ANSI CSI escape sequence (not just color/SGR codes)
function Get-VisibleLength ($text) {
    if ([string]::IsNullOrEmpty($text)) { return 0 }
    return ([regex]::Replace($text, "$ESC\[[0-9;]*[a-zA-Z]", "")).Length
}

# Dedicated adaptive-layout engine: measures the true visible width of each whole element
# ("label: value", color tags excluded from the count) and prints them side-by-side until
# they no longer fit the console width, then wraps to a new line. Elements are always kept
# whole and are never split across lines. Empty/null elements are skipped so they never
# introduce a stray separator. Used for both the portfolio metrics rows and the individual
# asset statistics rows.
function Get-WrappedSegmentLines {
    param(
        [string[]]$Segments,
        [int]$ConsoleWidth,
        [string]$Separator = "  ",
        [string]$ContinuationIndent = "  "
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $currentLine = $null
    $currentLen = 0
    $sepLen = Get-VisibleLength $Separator
    $indentLen = Get-VisibleLength $ContinuationIndent

    foreach ($seg in $Segments) {
        if ([string]::IsNullOrEmpty($seg)) { continue }

        $segLen = Get-VisibleLength $seg
        if ($null -eq $currentLine) {
            $currentLine = $seg
            $currentLen = $segLen
            continue
        }

        if (($currentLen + $sepLen + $segLen) -le $ConsoleWidth) {
            $currentLine += $Separator + $seg
            $currentLen += $sepLen + $segLen
        } else {
            $lines.Add($currentLine)
            $currentLine = $ContinuationIndent + $seg
            $currentLen = $indentLen + $segLen
        }
    }

    if ($null -ne $currentLine) {
        $lines.Add($currentLine)
    }

    return , $lines.ToArray()
}

# Same wrapping algorithm as Get-WrappedSegmentLines, but instead of returning the assembled
# lines it returns which segment indices are forced onto a new line. Used to work out, for a
# single asset, where IT would break - so those per-asset results can then be combined (OR'd)
# into one common set of break points shared by every asset in the table (see Get-LinesFromBreaks).
function Get-SegmentBreaksBefore {
    param(
        [string[]]$Segments,
        [int]$ConsoleWidth,
        [string]$Separator = "  ",
        [string]$ContinuationIndent = "  "
    )

    $breaksBefore = New-Object bool[] ($Segments.Count)
    $sepLen = Get-VisibleLength $Separator
    $indentLen = Get-VisibleLength $ContinuationIndent
    $currentLen = -1

    for ($i = 0; $i -lt $Segments.Count; $i++) {
        $seg = $Segments[$i]
        if ([string]::IsNullOrEmpty($seg)) { continue }

        $segLen = Get-VisibleLength $seg
        if ($currentLen -lt 0) {
            $currentLen = $segLen
            continue
        }

        if (($currentLen + $sepLen + $segLen) -le $ConsoleWidth) {
            $currentLen += $sepLen + $segLen
        } else {
            $breaksBefore[$i] = $true
            $currentLen = $indentLen + $segLen
        }
    }

    return , $breaksBefore
}

# Assembles Segments into lines using a fixed, precomputed set of break points (as returned/
# combined from Get-SegmentBreaksBefore) instead of measuring widths itself. This is what lets
# every asset row in the table break at the same segment: the break points are decided once
# (from the widest-wrapping asset) and then simply replayed here for every asset, regardless of
# how narrow that particular asset's own content is.
function Get-LinesFromBreaks {
    param(
        [string[]]$Segments,
        [bool[]]$BreaksBefore,
        [string]$Separator = "  ",
        [string]$ContinuationIndent = "  "
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $currentLine = $null

    for ($i = 0; $i -lt $Segments.Count; $i++) {
        $seg = $Segments[$i]
        if ([string]::IsNullOrEmpty($seg)) { continue }

        if ($null -eq $currentLine) {
            $currentLine = $seg
            continue
        }

        if ($i -lt $BreaksBefore.Count -and $BreaksBefore[$i]) {
            $lines.Add($currentLine)
            $currentLine = $ContinuationIndent + $seg
        } else {
            $currentLine += $Separator + $seg
        }
    }

    if ($null -ne $currentLine) {
        $lines.Add($currentLine)
    }

    return , $lines.ToArray()
}

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
        Enabled = $true
        GraphWidth = $script:AssetGraphWidth
        GraphGap = $script:AssetGraphGap
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
    } else {
        Get-WrappedSegmentLines -Segments $Segments -ConsoleWidth $textAreaWidth -Separator $Separator -ContinuationIndent $ContinuationIndent
    }
    if ($lines.Count -eq 0) { $lines = @("") }

    $renderRowCount = if ($graphEnabled) { [math]::Max($lines.Count, $GraphRows.Count) } else { $lines.Count }

    for ($i = 0; $i -lt $renderRowCount; $i++) {
        $textLine = if ($i -lt $lines.Count) { $lines[$i] } else { "" }
        $graphRow = if ($graphEnabled -and $i -lt $GraphRows.Count) { $GraphRows[$i] } else { [string]([char]0x2800) * $GraphWidth }

        if ($graphEnabled) {
            $textVisLen = Get-VisibleLength $textLine
            $padToGraphColumn = [math]::Max($GraphGap, $GraphColumnStart - $textVisLen)
            $coloredGraph = "${GraphColor}${graphRow}${Reset}"
            $null = $OutputBuffer.AppendLine("$textLine$(' ' * $padToGraphColumn)$coloredGraph")
        } else {
            $null = $OutputBuffer.AppendLine($textLine)
        }
    }
}

# --- PORTFOLIO METRICS SUMMARY STRINGS (computed here, before the chart-height sizing
# below, so Get-WrappedSegmentRowCount can measure their real wrapped row count instead of
# the sizing logic having to guess it from console-width bands alone) ---
if ($csvMetricsAvailable) {

    $totalGainAllTime = $totalValue + $totalWithdrawn - $totalDeposited
    if ($totalDeposited -ne 0) {
        $costRatio = if ($totalGainAllTime -gt 0) { ($totalExpensesSustained / $totalRealizedGain) * 100 } else { $null }
    } elseif ($totalCost -ne 0) {
        $costRatio = ($totalExpensesSustained / $totalCost) * 100
    } else {
        $costRatio = $null
    }
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

# Determine maximum asset description length for adaptive single-row layout on wide consoles
$maxDescLen = 0
foreach ($item in $portfolio) {
    $descLen = ("$($item.Ticker) $($item.Name)").Length
    if ($descLen -gt $maxDescLen) { $maxDescLen = $descLen }
}

# FIX: the single-row-per-asset layout used to switch on at a fixed "$consoleWidth -ge 135"
# threshold, regardless of how much horizontal space the description/metrics/P/L/sparkline
# actually needed. That guess was frequently wrong in both directions: on a console just
# past 135 columns wide (but with long asset names) the combined "desc+metrics" segment and
# the separate "P/L" segment did not actually fit together on the text area left after the
# sparkline's worst-case ~28-column reservation, so Get-WrappedSegmentLines still split them
# onto two lines - i.e. "single row" mode rendered two rows anyway. Conversely, on consoles
# with short asset names the true single-row content could fit well below 135 too, wasting
# the more compact single-row layout on consoles that could support it.
# Fix: measure the ACTUAL fixed-width pieces (description, padding, metrics columns, P/L
# columns, sparkline reservation) using the very same format strings used to render them
# further down, and only switch to single-row mode when the full assembled row is guaranteed
# to fit within $consoleWidth - keeping the "row count" decision in sync with reality instead
# of a magic number. This threshold check intentionally stays conservative (worst-case 27-wide
# graph reservation): it only decides a yes/no layout switch, not how many lines get printed.
$sampleMetricsPart = "${Gray}Qty:${Reset} {0,4}  ${Gray}Avg:${Reset} {1,8}  ${Gray}Price:${Reset} {2,8}  ${Gray}Value:${Reset} ${White}{3,9} ({4,5}%)${Reset}" -f "0", "0", "0", "0", "0"
$sampleGainText = "{0} {1,9} ({2,6}%)" -f ([char]0x2192), "0", "0"
$samplePlPart = "P/L: $sampleGainText   Day: $sampleGainText"

$singleRowLabelAndDescWidth   = 12 + $maxDescLen                    # category label (11) + 1 space + "Ticker Name"
$singleRowPaddingWidth        = 2                                   # minimum padding between description and metrics columns
$singleRowMetricsWidth        = Get-VisibleLength $sampleMetricsPart
$singleRowSegmentSeparatorWidth = 3                                 # "   " separator Get-WrappedSegmentLines inserts between same-line segments
$singleRowPlWidth             = Get-VisibleLength $samplePlPart
$singleRowGraphReserve        = if ($assetGraphLayoutProbe.Enabled) { 9 + $script:AssetGraphGap } else { 0 }

$requiredSingleRowWidth = $singleRowLabelAndDescWidth + $singleRowPaddingWidth + $singleRowMetricsWidth + $singleRowSegmentSeparatorWidth + $singleRowPlWidth + $singleRowGraphReserve

$useSingleRowAssetLayout = $consoleWidth -ge $requiredSingleRowWidth

$maxSingleSegmentLen = 0

# --- Pass 1: build every asset's Segments (text only - no wrapping/graph sizing yet). ---
$assetSegmentsByItem = @()

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

    if ($useSingleRowAssetLayout) {
        # Single row per asset format for very wide consoles, aligned with max description length
        $descPart = "{0} {1} {2}" -f $labelFormatted, "${White}$($item.Ticker)${Reset}", "${Reset}$($item.Name)${Reset}"

        $paddingNeeded = [math]::Max(2, $maxDescLen - ("$($item.Ticker) $($item.Name)").Length + 2)
        $paddingStr = " " * $paddingNeeded

        $metricsPart = "${Gray}Qty:${Reset} {0,4}  ${Gray}Avg:${Reset} {1,8}  ${Gray}Price:${Reset} {2,8}  ${Gray}Value:${Reset} ${White}{3,9} ({4,5}%)${Reset}" -f `
            $qtyStr, $avgCostStr, $priceStr, $valStr, $weightStr

        $rawGainText = "{0} {1,9} ({2,6}%)" -f (Get-TrendArrow $item.Gain), $gainStr, $gainPctStr
        $rawDayText  = "{0} {1,8} ({2,6}%)" -f (Get-TrendArrow $item.DayChangeTotal), $dayStr, $dayPctStr
        $coloredGain = "${gainColor}${rawGainText}${Reset}"
        $coloredDay  = "${dcColor}${rawDayText}${Reset}"
        $plPart = "${Gray}P/L:${Reset} $coloredGain   ${Gray}Day:${Reset} $coloredDay"

        $assetSegments = @("$descPart$paddingStr$metricsPart", $plPart)
    } else {
        $titleLine = "$labelFormatted ${White}$($item.Ticker) ${Reset}$($item.Name)${Reset}"

        $posRow = " ${Gray}Qty:${Reset} {0,4}   ${Gray}Avg:${Reset} {1,8}   ${Gray}Price:${Reset} {2,8}   ${Gray}Value: ${White}{3,9} ({4,5}%)${Reset}" -f `
            $qtyStr, $avgCostStr, $priceStr, $valStr, $weightStr

        $rawGainText = "{0} {1,9} ({2,6}%)" -f (Get-TrendArrow $item.Gain), $gainStr, $gainPctStr
        $rawDayText  = "{0} {1,8} ({2,6}%)" -f (Get-TrendArrow $item.DayChangeTotal), $dayStr, $dayPctStr
        $gainFormatted = "{0,-21}" -f $rawGainText
        $dayFormatted  = "{0,-21}" -f $rawDayText
        $coloredGain = "${gainColor}${gainFormatted}${Reset}"
        $coloredDay  = "${dcColor}${dayFormatted}${Reset}"
        $plString = "${Gray}P/L:${Reset} ${coloredGain}   ${Gray}Day:${Reset} $coloredDay"

        $assetSegments = @($titleLine, $posRow, $plString)
    }

    foreach ($seg in $assetSegments) {
        $segLen = Get-VisibleLength $seg
        if ($segLen -gt $maxSingleSegmentLen) { $maxSingleSegmentLen = $segLen }
    }

    $assetSegmentsByItem += [pscustomobject]@{
        Segments = $assetSegments
        GraphGap = $assetGraphGap
        GraphColor = $graphColor
        Item = $item
    }
}

$assetSegmentCount = if ($assetSegmentsByItem.Count -gt 0) { $assetSegmentsByItem[0].Segments.Count } else { 0 }

# Builds one shared set of forced line-break points for the WHOLE table at a given text
# width: each asset's own break points (Get-SegmentBreaksBefore) are OR'd together, so the
# table breaks wherever the widest-wrapping asset needs to - and every row, even much
# narrower ones, is forced to break at that same point too. That's what keeps the
# Qty/Avg/Price/P/L columns vertically aligned across the whole table.
# ForceBreakBefore additionally marks segment indices that must ALWAYS start a new line,
# regardless of whether they'd technically fit on the previous one - see its use below.
function Get-GlobalBreaksBefore {
    param(
        [object[]]$Entries,
        [int]$TextWidth,
        [int]$SegmentCount,
        [int[]]$ForceBreakBefore = @()
    )
    $breaks = New-Object bool[] ($SegmentCount)
    foreach ($entry in $Entries) {
        $itemBreaksBefore = Get-SegmentBreaksBefore -Segments $entry.Segments -ConsoleWidth $TextWidth -Separator "   " -ContinuationIndent "  "
        for ($breakIdx = 0; $breakIdx -lt $itemBreaksBefore.Count; $breakIdx++) {
            if ($itemBreaksBefore[$breakIdx]) { $breaks[$breakIdx] = $true }
        }
    }
    foreach ($forceIdx in $ForceBreakBefore) {
        if ($forceIdx -ge 0 -and $forceIdx -lt $breaks.Count) { $breaks[$forceIdx] = $true }
    }
    return , $breaks
}

# In the multi-row layout, Segments are @(titleLine, posRow, plString). titleLine's length
# varies per asset (it contains the ticker/name), while posRow and plString are built entirely
# from fixed-width format specifiers ({0,4}, {0,-21}, etc.) and are therefore the SAME visible
# length for every asset. Packing titleLine together with posRow onto one line - which the
# width-based wrapping above would otherwise happily do once the console is wide enough - was
# what broke the vertical alignment: posRow's "Qty:"/"Avg:"/"Price:" labels would then start at
# a column that shifts with every asset's ticker/name length instead of a fixed column. Forcing
# a break right after titleLine keeps every OTHER line built purely from fixed-width segments,
# so whenever those lines combine (or not) they still start at the same column for every asset.
$forceBreakBefore = if (-not $useSingleRowAssetLayout -and $assetSegmentCount -ge 2) { @(1) } else { @() }

# --- Pass 1b: decide how many lines the asset rows will occupy, THEN size the sparkline to
# match - not the other way round. For a fixed-height graph (consoleWidth 78-136, always 3
# rows) the reserved width is exactly known up front, so this is a single pass. For the
# "dynamic" case (very wide/narrow consoles) the graph's own height - and therefore how much
# text width is left before the table wraps - isn't known until the line count is decided,
# which is itself decided by how much text width is left: a circular dependency. That's
# resolved by iterating: start from the shortest possible graph (1 row), see how many lines
# that implies for the widest-wrapping asset, and if that's more rows than assumed, grow the
# graph to match and re-check. This converges in at most 3 steps, since height only ever takes
# the values 1, 2 or 3 - and it never assumes a wider reservation than the table actually ends
# up needing, so assets no longer wrap onto an extra line (or get an over-tall sparkline) just
# because a worst-case guess reserved more room than was really required.
if ($assetGraphLayoutProbe.Enabled -and $null -eq $assetGraphLayoutProbe.GraphHeight) {
    $globalGraphHeight = 1
    for ($iter = 0; $iter -lt 3; $iter++) {
        $reserveWidth = [math]::Min(27, [math]::Max(9, 9 * $globalGraphHeight)) + $assetGraphGap
        $textWidth = [math]::Max(1, $consoleWidth - $reserveWidth)

        $globalBreaksBefore = Get-GlobalBreaksBefore -Entries $assetSegmentsByItem -TextWidth $textWidth -SegmentCount $assetSegmentCount -ForceBreakBefore $forceBreakBefore
        $lineCount = 1 + (@($globalBreaksBefore) | Where-Object { $_ }).Count
        $neededHeight = [math]::Min(3, [math]::Max(1, $lineCount))

        if ($neededHeight -le $globalGraphHeight) { break }
        $globalGraphHeight = $neededHeight
    }
} elseif ($assetGraphLayoutProbe.Enabled) {
    # Fixed height (always 3 rows -> always a 27-wide reservation): the reserved width is
    # already exact, no guessing/iterating needed.
    $globalGraphHeight = 3
    $textWidth = [math]::Max(1, $consoleWidth - 27 - $assetGraphGap)
    $globalBreaksBefore = Get-GlobalBreaksBefore -Entries $assetSegmentsByItem -TextWidth $textWidth -SegmentCount $assetSegmentCount -ForceBreakBefore $forceBreakBefore
} else {
    # No sparkline at all: wrap against the full console width.
    $globalGraphHeight = 0
    $globalBreaksBefore = Get-GlobalBreaksBefore -Entries $assetSegmentsByItem -TextWidth $consoleWidth -SegmentCount $assetSegmentCount -ForceBreakBefore $forceBreakBefore
}

$globalGraphWidth = if ($assetGraphLayoutProbe.Enabled) { [math]::Min(27, [math]::Max(9, 9 * $globalGraphHeight)) } else { 0 }

# --- Pass 2: replay the SAME break points for every asset (Get-LinesFromBreaks), so all rows
# share an identical line structure, using the sparkline height/width already decided above. ---
$assetRenderInfos = @()

foreach ($entry in $assetSegmentsByItem) {
    $wrappedTextLines = Get-LinesFromBreaks -Segments $entry.Segments -BreaksBefore $globalBreaksBefore -Separator "   " -ContinuationIndent "  "
    if ($wrappedTextLines.Count -eq 0) { $wrappedTextLines = @("") }

    $assetRenderInfos += [pscustomobject]@{
        Segments = $entry.Segments
        GraphWidth = $globalGraphWidth
        GraphGap = $entry.GraphGap
        GraphHeight = $globalGraphHeight
        WrappedTextLines = $wrappedTextLines
        GraphColor = $entry.GraphColor
        Item = $entry.Item
    }
}

$maxTextWidth = 0
foreach ($renderInfo in $assetRenderInfos) {
    foreach ($line in $renderInfo.WrappedTextLines) {
        $lineLen = Get-VisibleLength $line
        if ($lineLen -gt $maxTextWidth) { $maxTextWidth = $lineLen }
    }
}
$graphColumnStart = if ($maxTextWidth -gt 0) { $maxTextWidth + 1 } else { 0 }

# Clamp the graph column so that text + gap + graph never exceeds $consoleWidth: without
# this, some combinations of console width and asset description length pushed the
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

# Whole segments (e.g. a wide Qty/Avg metrics row) are never split, so the graph column
# can never start before the widest single segment ends - otherwise that segment alone
# would overflow into (or past) the graph, wrapping the terminal mid-line. If honoring
# that minimum would leave no usable room for even the narrowest sparkline, drop the
# per-asset graphs entirely for this render instead of printing a garbled line.
$minRequiredGraphColumnStart = $maxSingleSegmentLen + 1
if ($maxGraphWidthUsed -gt 0 -and ($minRequiredGraphColumnStart + $script:AssetGraphGap + $script:AssetGraphWidth) -gt $consoleWidth) {
    foreach ($renderInfo in $assetRenderInfos) { $renderInfo.GraphWidth = 0 }
    $graphColumnStart = 0
} elseif ($graphColumnStart -lt $minRequiredGraphColumnStart) {
    $graphColumnStart = $minRequiredGraphColumnStart
    $maxAllowedGraphWidth = [math]::Max(0, $consoleWidth - $graphColumnStart - $script:AssetGraphGap)
    foreach ($renderInfo in $assetRenderInfos) {
        if ($renderInfo.GraphWidth -gt $maxAllowedGraphWidth) { $renderInfo.GraphWidth = $maxAllowedGraphWidth }
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
# Assets table block: use the actual final row count from the shared break points.
$assetRowsPerItem = if ($assetRenderInfos.Count -gt 0) {
    [math]::Max(1, @($assetRenderInfos[0].WrappedTextLines).Count)
} else {
    0
}
if ($assets.Count -gt 0) {
    $fixedRows += $assets.Count * ($assetRowsPerItem + 1)
}
# Some space for elaborated PS Prompts at the bottom of the screen:
$fixedRows += 2
# When launched from the Explorer RMB context menu, reserve 1 extra line for the
# "Press any key to close..." prompt printed at the very end of the script.
if ($script:launchedFromExplorer) { $fixedRows += 1 }

$availableHeightForCharts = $consoleHeight - $fixedRows
$calculatedChartHeight = [math]::Max($minChartHeight, [int][math]::Floor($availableHeightForCharts / 2))

# Chart Generation with Dynamic Height
$intradayChartRows = Get-BrailleSparkline -values $portfolioIntradayFiltered -height $calculatedChartHeight -targetWidth $targetChartWidth -visibleRatio $intradayRatio
$intradayTimelineRow = Get-IntradayTimelineRow -targetWidth $targetChartWidth

$customChartRows = Get-BrailleSparkline -values $portfolioCustom -height $calculatedChartHeight -targetWidth $targetChartWidth -Filled
$customTimelineRow = Get-YearlyTimelineRow -timestamps $customTimestamps -targetWidth $targetChartWidth

$outputBuffer = [System.Text.StringBuilder]::new()

$null = $outputBuffer.Append("$ESC[2J$ESC[H")

# A. INTRADAY Chart Block
$null = $outputBuffer.AppendLine("${White}Portfolio performance $((Get-Date).ToString('dd/MM/yyyy HH:mm'))${Reset}")

$intraFirstVal = $portfolioPrevCloseValue
$intraLastVal = if ($portfolioIntradayFiltered.Count -gt 0) { $portfolioIntradayFiltered[-1] } else { 0 }
$intraChartColor = Get-TrendColor ($intraLastVal - $intraFirstVal)

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
} else {
    $twrrStr = ""
}

if ($csvMetricsAvailable -and $null -ne $mwrr) {
    $mwrrColor = Get-TrendColor $mwrr
    $mwrrPctVal = $mwrr * 100
    $mwrrSign = if ($mwrrPctVal -gt 0) { "+" } else { "" }
    $mwrrStr = "MWRR: ${mwrrColor}${mwrrSign}$(Format-NumberLocalized $mwrrPctVal 2)% ${Reset}annual"
} else {
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

    if (-not (Test-Path $logPath)) {
        $headerLine = (@("Timestamp", "LogLevel") + @($logStats.Keys)) -join "`t"
        Set-Content -Path $logPath -Value $headerLine -Encoding UTF8
    }

    $logTimestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $logLine = (@($logTimestamp, "INFO") + @($logStats.Values)) -join "`t"
    Add-Content -Path $logPath -Value $logLine -Encoding UTF8
} catch {
    Write-Warning "WARNING: Unable to write log row to ${logPath}: $_"
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

    Write-WrappedSegmentsWithRightGraph -OutputBuffer $outputBuffer -Segments $renderInfo.Segments -ConsoleWidth $consoleWidth -Separator "   " -ContinuationIndent "  " -GraphRows $graphRows -GraphWidth $renderInfo.GraphWidth -GraphGap $renderInfo.GraphGap -GraphColumnStart $graphColumnStart -GraphColor $renderInfo.GraphColor -Reset $Reset -BreaksBefore $globalBreaksBefore

    $null = $outputBuffer.AppendLine("")
}

Write-Host $outputBuffer.ToString() -NoNewline

# When launched from the Explorer RMB context menu, the console window would otherwise
# close immediately once the script returns, so keep it open until the user acknowledges.
if ($script:launchedFromExplorer) {
    Write-Host "${Gray}Press any key to close...${Reset}" -NoNewline
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
